import os
import sys
import torch
import argparse
from flask import Flask, request, jsonify
from transformers import AutoModelForCausalLM, AutoTokenizer, BitsAndBytesConfig
import preprocess
import traceback
from torch import nn
import math
from peft import LoraConfig, PeftModel

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# Dynamic positional encoding
class DynamicPositionalEncoding(nn.Module):
    def __init__(self, d_model, max_len=512):
        super(DynamicPositionalEncoding, self).__init__()
        self.d_model = d_model
        self.max_len = max_len
        
        # Create a constant positional encoding matrix
        pe = torch.zeros(max_len, d_model)
        position = torch.arange(0, max_len, dtype=torch.float).unsqueeze(1)
        div_term = torch.exp(torch.arange(0, d_model, 2).float() * (-math.log(10000.0) / d_model))
        pe[:, 0::2] = torch.sin(position * div_term)
        pe[:, 1::2] = torch.cos(position * div_term)
        pe = pe.unsqueeze(0)  # Shape: [1, max_len, d_model]
        self.register_buffer('pe', pe)

    def forward(self):
        return self.pe

# Protein structure sequence adapter
class ProteinStructureSequenceAdapter(nn.Module):
    def __init__(self, input_dim, output_dim, num_heads, num_queries, max_len=512):
        super(ProteinStructureSequenceAdapter, self).__init__()
        self.input_dim = input_dim
        self.output_dim = output_dim
        self.num_heads = num_heads
        self.num_queries = num_queries
        self.max_len = max_len
        
        # Linear projection layer
        self.linear_proj = nn.Linear(input_dim, output_dim)
        
        # Dynamic positional encoding layer
        self.pos_encoder = DynamicPositionalEncoding(output_dim, max_len)
        
        # Learnable queries
        self.learnable_queries = nn.Parameter(torch.randn(num_queries, output_dim))
        
        # Cross-attention layer
        self.cross_attention = nn.MultiheadAttention(embed_dim=output_dim, num_heads=num_heads, batch_first=True)
        
        # Output projection layer
        self.output_proj = nn.Linear(output_dim, output_dim)
        
        self.question_proj = nn.Linear(output_dim, output_dim)
        # Layer normalization
        self.layer_norm1 = nn.LayerNorm(output_dim)
        self.layer_norm2 = nn.LayerNorm(output_dim)
        self.layer_norm3 = nn.LayerNorm(output_dim)

    def forward(self, x, h_state):
        # x: [batch, seq_len, input_dim]
        # h_state: [batch, hidden_dim] — single question vector, matching training
        batch_size, seq_len, _ = x.size()

        if seq_len < self.max_len:
            pad_size = self.max_len - seq_len
            padding = torch.zeros(batch_size, pad_size, self.input_dim, device=x.device, dtype=x.dtype)
            x = torch.cat([x, padding], dim=1)
        elif seq_len > self.max_len:
            x = x[:, :self.max_len, :]

        x_proj = self.linear_proj(x)
        x_proj = self.layer_norm1(x_proj)
        pe = self.pos_encoder()
        x_pos_encoded = x_proj + pe[:, :x_proj.size(1), :]

        queries = self.learnable_queries.unsqueeze(0).expand(batch_size, -1, -1)
        # question_proj expects [batch, hidden_dim]; unsqueeze(1) broadcasts to all queries
        h_state = self.question_proj(h_state).unsqueeze(1)
        queries = queries + h_state
        queries = self.layer_norm2(queries)
        queries_pos_encoded = pe[:, :self.num_queries*2:2, :] + queries

        attn_output, _ = self.cross_attention(queries_pos_encoded, x_pos_encoded, x_pos_encoded)
        attn_output = self.layer_norm3(attn_output)
        output = self.output_proj(attn_output)
        return output

# Global variables to store the loaded model and adapter
model = None
tokenizer = None
adapter = None
device = None
_zero_question = False

# Initialize the model and adapter
def initialize_models(model_path, lora_path, adapter_path, skip_lora=False, no_quant=False, zero_question=False):
    global model, tokenizer, adapter, device, _zero_question
    _zero_question = zero_question

    # Set the device
    device = "cuda" if torch.cuda.is_available() else "cpu"
    print(f"Using device: {device}")

    # Load model L (LLaMA model)
    print(f"Loading model: {model_path}")
    if torch.cuda.is_available():
        if no_quant:
            # float16 with device_map=auto so it can offload layers to CPU RAM
            # if GPU VRAM is insufficient (e.g. 6 GB). Eliminates quantization drift.
            print("Loading model in float16 (no quantization, auto device map)")
            model = AutoModelForCausalLM.from_pretrained(
                model_path,
                torch_dtype=torch.float16,
                device_map="auto",
                low_cpu_mem_usage=True,
            )
        else:
            bnb_config = BitsAndBytesConfig(
                load_in_4bit=True,
                # bfloat16 matches the model's native training dtype (LLaMA 3 is bf16).
                # Using float16 here caused hidden-state drift that fed into the adapter.
                bnb_4bit_compute_dtype=torch.bfloat16,
                bnb_4bit_use_double_quant=True,
                bnb_4bit_quant_type="nf4",
            )
            # BitsAndBytes 4-bit cannot offload layers to CPU — must stay on one GPU.
            model = AutoModelForCausalLM.from_pretrained(
                model_path,
                quantization_config=bnb_config,
                device_map="cuda:0",
                low_cpu_mem_usage=True,
            )
        torch.cuda.empty_cache()
    else:
        model = AutoModelForCausalLM.from_pretrained(
            model_path,
            torch_dtype=torch.float16,
            device_map="cpu",
        )

    # Load LoRA weights
    if skip_lora:
        print("Skipping LoRA weights (--no_lora flag set)")
        model.eval()
    else:
        print(f"Loading LoRA weights: {lora_path}")
        model = PeftModel.from_pretrained(model, model_id=lora_path)
        model.eval()

    # Load tokenizer — prefer lora_path dir (has tokenizer.json), fall back to model_path
    print("Loading tokenizer")
    tokenizer_path = lora_path if os.path.isfile(os.path.join(lora_path, "tokenizer.json")) else model_path
    tokenizer = AutoTokenizer.from_pretrained(tokenizer_path)
    if tokenizer.pad_token is None:
        tokenizer.pad_token = tokenizer.eos_token

    # Load adapter model A
    print(f"Loading adapter model: {adapter_path}")
    adapter = ProteinStructureSequenceAdapter(input_dim=1152, output_dim=4096, num_heads=16, num_queries=256, max_len=512)
    checkpoint = torch.load(adapter_path, map_location="cpu", weights_only=False)
    adapter.load_state_dict(checkpoint['adapter_model_weight'])
    adapter = adapter.to(device).half()  # float16 — matches training's autocast float16
    adapter.eval()
    
    if torch.cuda.is_available():
        allocated = torch.cuda.memory_allocated() / 1024**3
        reserved = torch.cuda.memory_reserved() / 1024**3
        print(f"GPU VRAM used: {allocated:.2f} GB allocated, {reserved:.2f} GB reserved")
    print("All models loaded successfully")

# Main function: Process PDB file and question, generate answer
def generate_answer(pdb_file_path, question):
    global model, tokenizer, adapter, device

    print(f"\n{'='*60}")
    print(f"[INFERENCE] PDB: {pdb_file_path}")
    print(f"[INFERENCE] Question: {question}")
    print(f"[INFERENCE] zero_question={_zero_question}")
    print(f"{'='*60}")

    # Ensure the models are loaded
    if model is None or tokenizer is None or adapter is None:
        return "Error: Models are not initialized"

    try:
        # Step 1: Process the PDB file → protein embedding (1152-dim per residue)
        print("[STEP 1] Running ProteinMPNN ensemble on PDB...")
        protein_vector = preprocess.get_mpnn_emb(pdb_file_path).unsqueeze(0).to(device)
        print(f"[STEP 1] Raw protein_vector shape: {protein_vector.shape}  "
              f"dtype: {protein_vector.dtype}  "
              f"norm(mean): {protein_vector.norm(dim=-1).mean().item():.4f}")

        max_length = 512
        orig_len = protein_vector.size(1)
        if max_length > orig_len:
            padding_length = max_length - orig_len
            padding = torch.zeros((1, padding_length, 1152), device=device)
            protein_vector = torch.cat([protein_vector, padding], dim=1)
            print(f"[STEP 1] Padded from {orig_len} → {max_length} residues")
        else:
            protein_vector = protein_vector[:, :max_length, :]
            print(f"[STEP 1] Truncated from {orig_len} → {max_length} residues")

        # Step 2: Get question hidden state Qht (paper eq. 8)
        inputs = tokenizer(f"Human: {question}\nAssistant: ", return_tensors="pt").to(device)
        print(f"[STEP 2] Tokenized question: {inputs.input_ids.shape[1]} tokens")

        if _zero_question:
            # 4-bit quantization distorts Qht after 32 layers; zeros let the adapter
            # rely on its trained learnable queries instead of corrupted hidden states.
            question_hidden_state = torch.zeros(1, 4096, device=device, dtype=torch.float16)
            print("[STEP 2] question_hidden_state: zeros (--zero_question)")
        else:
            print("[STEP 2] Running LLM forward pass for Qht...")
            with torch.no_grad():
                hidden_states = model(
                    inputs.input_ids,
                    attention_mask=inputs.attention_mask,
                    output_hidden_states=True
                ).hidden_states[-1]
            question_hidden_state = hidden_states[:, -1, :].to(torch.float16)
            qh_norm = question_hidden_state.norm().item()
            print(f"[STEP 2] question_hidden_state norm: {qh_norm:.4f}  "
                  f"min: {question_hidden_state.min().item():.4f}  "
                  f"max: {question_hidden_state.max().item():.4f}")

        # Step 3: Run adapter (protein_vector + Qht → 256 soft-prompt tokens)
        print("[STEP 3] Running adapter...")
        protein_vector = protein_vector.half()
        with torch.no_grad():
            protein_embedding = adapter(protein_vector, question_hidden_state)
        print(f"[STEP 3] protein_embedding shape: {protein_embedding.shape}  "
              f"dtype: {protein_embedding.dtype}  "
              f"norm(mean): {protein_embedding.norm(dim=-1).mean().item():.4f}  "
              f"min: {protein_embedding.min().item():.4f}  "
              f"max: {protein_embedding.max().item():.4f}")

        # Check for NaN/Inf — if present, adapter received bad input
        if torch.isnan(protein_embedding).any():
            print("[STEP 3] WARNING: protein_embedding contains NaN values!")
        if torch.isinf(protein_embedding).any():
            print("[STEP 3] WARNING: protein_embedding contains Inf values!")

        # Step 4: Encode question text into token embeddings
        # PeftModel: model.base_model.model.model.embed_tokens
        # plain LlamaForCausalLM: model.model.embed_tokens
        try:
            embed_tokens = model.base_model.model.model.embed_tokens
        except AttributeError:
            embed_tokens = model.model.embed_tokens
        inputs_embeds = embed_tokens(inputs.input_ids)
        protein_embedding = protein_embedding.to(dtype=inputs_embeds.dtype)

        prot_norm = protein_embedding.norm(dim=-1).mean().item()
        text_norm = inputs_embeds.norm(dim=-1).mean().item()
        print(f"[STEP 4] text inputs_embeds shape: {inputs_embeds.shape}  "
              f"norm(mean): {text_norm:.4f}")
        print(f"[DIAG]   protein_embedding norm : {prot_norm:.4f}")
        print(f"[DIAG]   text_embeds norm       : {text_norm:.4f}")
        print(f"[DIAG]   norm ratio (prot/text) : {prot_norm / (text_norm + 1e-8):.4f}  "
              f"(ideally close to 1.0)")

        # Step 5: Prepend protein soft-prompt to question embeddings
        combined_embeds = torch.cat([protein_embedding, inputs_embeds], dim=1)
        combined_attention_mask = torch.cat([
            torch.ones((protein_embedding.size(0), protein_embedding.size(1)),
                       device=device, dtype=inputs.attention_mask.dtype),
            inputs.attention_mask
        ], dim=1)
        print(f"[STEP 5] combined_embeds shape: {combined_embeds.shape}")

        if torch.cuda.is_available():
            alloc = torch.cuda.memory_allocated() / 1024**3
            reserv = torch.cuda.memory_reserved() / 1024**3
            print(f"[STEP 5] GPU VRAM before generate: {alloc:.2f} GB alloc / {reserv:.2f} GB reserved")

        # Step 6: Generate
        print("[STEP 6] Generating response...")
        with torch.no_grad():
            generated_ids = model.generate(
                inputs_embeds=combined_embeds,
                attention_mask=combined_attention_mask,
                pad_token_id=tokenizer.eos_token_id,
                eos_token_id=tokenizer.eos_token_id,
                max_new_tokens=128,
                do_sample=True,       # sampling avoids greedy error lock-in under 4-bit quantization
                temperature=0.3,      # low temperature = focused but not deterministic
                top_p=0.9,
                repetition_penalty=1.3,
                no_repeat_ngram_size=4,
            )
        print(f"[STEP 6] Generated {generated_ids.shape[1]} tokens")

        response = tokenizer.decode(generated_ids[0], skip_special_tokens=True)
        print(f"[STEP 6] Raw decoded response: {repr(response[:200])}")

        if "Assistant:" in response:
            response = response.split("Assistant:")[1].strip()

        print(f"[INFERENCE] Final answer: {response[:300]}")
        print(f"{'='*60}\n")
        return response

    except Exception as e:
        print(f"[ERROR] generate_answer failed: {str(e)}")
        print(traceback.format_exc())
        return f"Error processing: {str(e)}"

# Create a Flask application
def create_app():
    app = Flask(__name__)
    
    # Create a simple HTML template
    @app.route('/')
    def index():
        return '''
        <!DOCTYPE html>
        <html>
        <head>
            <title>Prot2Chat: Protein Q&A System</title>
            <style>
                body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
                .form-group { margin-bottom: 15px; }
                label { display: block; margin-bottom: 5px; }
                input[type="file"], input[type="text"] { width: 100%; padding: 8px; }
                button { padding: 10px 15px; background-color: #4CAF50; color: white; border: none; cursor: pointer; }
                #result { margin-top: 20px; padding: 15px; border: 1px solid #ddd; border-radius: 4px; }
                #loading { display: none; margin-top: 20px; text-align: center; }
            </style>
        </head>
        <body>
            <h1>Prot2Chat: Protein Q&A System</h1>
            <form id="queryForm" enctype="multipart/form-data">
                <div class="form-group">
                    <label for="pdbFile">Upload PDB file:</label>
                    <input type="file" id="pdbFile" name="pdbFile" accept=".pdb" required>
                </div>
                <div class="form-group">
                    <label for="question">Enter your question:</label>
                    <textarea id="question" name="question" placeholder="Please enter your question about protein..." required style="width: 100%; height: 100px;"></textarea>
                </div>
                <button type="submit">Submit</button>
            </form>
            <div id="loading">Processing, please wait...</div>
            <div id="result" style="display: none;">
                <h2>Answer:</h2>
                <p id="answer"></p>
            </div>

            <script>
                document.getElementById('queryForm').addEventListener('submit', async function(e) {
                    e.preventDefault();
                    
                    const formData = new FormData();
                    formData.append('pdbFile', document.getElementById('pdbFile').files[0]);
                    formData.append('question', document.getElementById('question').value);
                    
                    document.getElementById('result').style.display = 'none';
                    document.getElementById('loading').style.display = 'block';
                    
                    try {
                        const response = await fetch('/api/query', {
                            method: 'POST',
                            body: formData
                        });
                        
                        const data = await response.json();
                        document.getElementById('loading').style.display = 'none';
                        
                        if (data.error) {
                            document.getElementById('answer').textContent = "Error: " + data.error;
                        } else {
                            document.getElementById('answer').textContent = data.answer;
                        }
                        document.getElementById('result').style.display = 'block';
                    } catch (error) {
                        document.getElementById('loading').style.display = 'none';
                        console.error('Error:', error);
                        document.getElementById('answer').textContent = "An error occurred, please try again: " + error;
                        document.getElementById('result').style.display = 'block';
                    }
                });
            </script>
        </body>
        </html>
        '''

    # API endpoint
    @app.route('/api/query', methods=['POST'])
    def query():
        try:
            # Get the uploaded PDB file
            if 'pdbFile' not in request.files:
                return jsonify({'error': 'No PDB file uploaded'}), 400
                
            pdb_file = request.files['pdbFile']
            if pdb_file.filename == '':
                return jsonify({'error': 'No file selected'}), 400
                
            question = request.form.get('question', '')
            if not question:
                return jsonify({'error': 'Question cannot be empty'}), 400
            
            # Save the PDB file to a temporary location
            temp_dir = os.path.join(os.getcwd(), 'temp')
            os.makedirs(temp_dir, exist_ok=True)
            temp_pdb_path = os.path.join(temp_dir, pdb_file.filename)
            pdb_file.save(temp_pdb_path)
            
            print(f"Processing file: {temp_pdb_path}")
            print(f"Question: {question}")
            
            # Generate the answer
            answer = generate_answer(temp_pdb_path, question)
            
            # Delete the temporary file
            if os.path.exists(temp_pdb_path):
                os.remove(temp_pdb_path)
            
            return jsonify({'answer': answer})
        except Exception as e:
            print(f"API error: {str(e)}")
            print(traceback.format_exc())
            return jsonify({'error': str(e)}), 500
    
    return app


if __name__ == '__main__':
    
    parser = argparse.ArgumentParser(description='Protein Structure Q&A System')
    parser.add_argument('--model_path', type=str,
                        default=os.path.join(_SCRIPT_DIR, 'base_model'),
                        help='Path to the LLaMA base model')
    parser.add_argument('--lora_path', type=str,
                        default=os.path.join(_SCRIPT_DIR, 'lora_wight'),
                        help='Path to the LoRA weights directory')
    parser.add_argument('--adapter_path', type=str,
                        default=os.path.join(_SCRIPT_DIR, 'adapter_weight', 'adapter_model_and_optimizer_1_400000.pth'),
                        help='Path to the adapter model weights (.pth)')
    parser.add_argument('--port', type=int, default=7777,
                        help='Server port')
    parser.add_argument('--gpu', type=str, default="",
                        help='IDs of the GPUs to use, e.g., "0,1". Leave empty for CPU.')
    parser.add_argument('--no_lora', action='store_true',
                        help='Skip loading LoRA weights (for diagnosing LoRA compatibility)')
    parser.add_argument('--no_quant', action='store_true',
                        help='Load model in float16 instead of 4-bit. Uses device_map=auto to offload layers to CPU when VRAM is limited. Matches training precision; recommended if you have >=16 GB system RAM.')
    parser.add_argument('--zero_question', action='store_true',
                        help='Use a zero vector for question conditioning instead of LLM hidden states. Bypasses 4-bit quantization distortion of Qht. Try this first on low-VRAM (<=8 GB) systems.')

    args = parser.parse_args()

    if args.gpu:
        os.environ["CUDA_VISIBLE_DEVICES"] = args.gpu

    initialize_models(args.model_path, args.lora_path, args.adapter_path,
                      skip_lora=args.no_lora, no_quant=args.no_quant,
                      zero_question=args.zero_question)
    
    
    app = create_app()
    app.run(host='localhost', port=args.port, debug=False)