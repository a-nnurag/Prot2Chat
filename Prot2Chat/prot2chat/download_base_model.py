"""
Downloads the Prot2Chat fine-tuned model from ModelScope into base_model_prot2chat/.
This is the exact model the adapter and LoRA weights were trained with.

Requirements:
  pip install modelscope

Usage:
  python download_base_model.py
"""
import os
from modelscope import snapshot_download

REPO_ID = "zcw51699/prot2chat"
DEST = os.path.join(os.path.dirname(os.path.abspath(__file__)), "base_model_prot2chat")

def main():
    print(f"Downloading {REPO_ID} → {DEST}")
    snapshot_download(model_id=REPO_ID, local_dir=DEST)
    print(f"Done. Model saved to {DEST}")
    print("Now run:")
    print(f"  python demo.py --model_path {DEST} --lora_path ./lora_wight --adapter_path ./adapter_weight/adapter_model_and_optimizer_1_400000.pth --port 7777 --gpu 0")

if __name__ == "__main__":
    main()
