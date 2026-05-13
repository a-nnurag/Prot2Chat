#!/usr/bin/env python3
"""
eval_proteintalk.py

Evaluates ProteinTalk on a sample of the Mol-Instructions protein test set.
Computes BLEU-2, ROUGE-1, ROUGE-2, ROUGE-L and compares with paper baselines.

Requirements:
    pip install datasets sacrebleu rouge-score requests

Usage:
    # Make sure ProteinTalk server is running first:
    #   python demo.py --adapter_path ./adapter_weight/adapter_model_and_optimizer_1_400000.pth --use_groq --port 7777 --gpu 0

    python eval_proteintalk.py --api_url http://localhost:7777 --n_samples 50
    python eval_proteintalk.py --api_url http://localhost:7777 --n_samples 20 --out results.json
"""

import os
import re
import sys
import json
import time
import random
import argparse
import tempfile
import requests

# ── Paper baseline scores (from Table 2 of btaf396.pdf, Mol-Instructions test set) ──
PAPER_SCORES = {
    "Prot2Chat (Wang et al.)": {"bleu2": 35.85, "rouge1": 57.21, "rouge2": 38.09, "rougeL": 50.51},
    "Evola-10B":               {"bleu2":  8.69, "rouge1": 29.09, "rouge2":  8.41, "rougeL": 20.04},
    "KIMI few-shot":           {"bleu2": 12.05, "rouge1": 31.21, "rouge2": 11.38, "rougeL": 24.18},
    "LLaMA3-FT":               {"bleu2":  6.42, "rouge1": 24.50, "rouge2":  6.32, "rougeL": 17.03},
    "BioMedGPT-LM-10B":        {"bleu2":  1.02, "rouge1": 10.93, "rouge2":  1.57, "rougeL":  7.84},
}


# ── Dataset loading ───────────────────────────────────────────────────────────────────

def load_test_samples(n_samples: int, seed: int = 42):
    """Load n_samples from Mol-Instructions protein test set."""
    print(f"[DATASET] Loading Mol-Instructions protein test set from HuggingFace...")
    from datasets import load_dataset
    ds = load_dataset("zjunlp/Mol-Instructions", "Protein-oriented instructions",
                      split="test", trust_remote_code=True)
    print(f"[DATASET] Total test samples: {len(ds)}")

    # Filter to protein function description tasks (have clear reference answers)
    samples = [s for s in ds if s.get("input", "").strip()]
    print(f"[DATASET] Samples with protein input: {len(samples)}")

    # Filter by sequence length for ESMFold compatibility (<=500 AA)
    aa_pattern = re.compile(r'^[ACDEFGHIKLMNPQRSTVWY]+$', re.IGNORECASE)
    valid = [s for s in samples
             if aa_pattern.match(s["input"].strip()) and len(s["input"].strip()) <= 500]
    print(f"[DATASET] Samples with valid AA sequence (<=500): {len(valid)}")

    random.seed(seed)
    selected = random.sample(valid, min(n_samples, len(valid)))
    print(f"[DATASET] Selected {len(selected)} samples for evaluation")
    return selected


# ── Structure prediction ──────────────────────────────────────────────────────────────

def fold_with_esmfold(sequence: str, out_path: str) -> bool:
    """Fold a protein sequence using ESMFold API → save PDB to out_path."""
    url = "https://api.esmatlas.com/foldSequence/v1/pdb/"
    try:
        resp = requests.post(url, data=sequence, timeout=120,
                             headers={"Content-Type": "application/x-www-form-urlencoded"})
        if resp.status_code == 200 and resp.text.startswith("ATOM"):
            with open(out_path, "w") as f:
                f.write(resp.text)
            return True
        print(f"  [ESMFold] Status {resp.status_code}: {resp.text[:100]}")
        return False
    except Exception as e:
        print(f"  [ESMFold] Error: {e}")
        return False


def fetch_alphafold_pdb(uniprot_id: str, out_path: str) -> bool:
    """Fetch AlphaFold predicted structure for a UniProt ID."""
    url = f"https://alphafold.ebi.ac.uk/files/AF-{uniprot_id}-F1-model_v4.pdb"
    try:
        resp = requests.get(url, timeout=30)
        if resp.status_code == 200:
            with open(out_path, "w") as f:
                f.write(resp.text)
            return True
        return False
    except Exception as e:
        print(f"  [AlphaFold] Error: {e}")
        return False


def sequence_to_uniprot(sequence: str) -> str | None:
    """Look up UniProt accession for an exact sequence match."""
    url = "https://rest.uniprot.org/uniprotkb/search"
    params = {"query": f'sequence:"{sequence}"', "fields": "accession", "format": "json", "size": 1}
    try:
        resp = requests.get(url, params=params, timeout=20)
        if resp.status_code == 200:
            data = resp.json()
            if data.get("results"):
                return data["results"][0]["primaryAccession"]
    except Exception:
        pass
    return None


def get_structure(sequence: str, tmp_dir: str, idx: int) -> str | None:
    """Get PDB file for a sequence. Tries ESMFold first, then AlphaFold via UniProt."""
    pdb_path = os.path.join(tmp_dir, f"protein_{idx}.pdb")

    print(f"  Folding with ESMFold (seq length={len(sequence)})...")
    if fold_with_esmfold(sequence, pdb_path):
        print(f"  ESMFold OK → {pdb_path}")
        return pdb_path

    print(f"  ESMFold failed. Trying AlphaFold via UniProt lookup...")
    uid = sequence_to_uniprot(sequence)
    if uid:
        print(f"  UniProt ID: {uid}")
        if fetch_alphafold_pdb(uid, pdb_path):
            print(f"  AlphaFold OK → {pdb_path}")
            return pdb_path

    print(f"  Could not get structure for sample {idx}")
    return None


# ── ProteinTalk API call ──────────────────────────────────────────────────────────────

def call_proteintalk(api_url: str, pdb_path: str, question: str) -> str:
    """Send PDB + question to the running ProteinTalk Flask API."""
    try:
        with open(pdb_path, "rb") as f:
            resp = requests.post(
                f"{api_url}/api/query",
                files={"pdbFile": (os.path.basename(pdb_path), f, "chemical/x-pdb")},
                data={"question": question},
                timeout=90,
            )
        if resp.status_code == 200:
            return resp.json().get("answer", "").strip()
        print(f"  [API] Error {resp.status_code}: {resp.text[:200]}")
        return ""
    except Exception as e:
        print(f"  [API] Exception: {e}")
        return ""


# ── Metrics ───────────────────────────────────────────────────────────────────────────

def compute_metrics(hypotheses: list[str], references: list[str]) -> dict:
    from sacrebleu.metrics import BLEU
    from rouge_score import rouge_scorer as rs

    # BLEU-2
    bleu = BLEU(max_ngram_order=2, effective_order=True)
    bleu_score = bleu.corpus_score(hypotheses, [references]).score

    # ROUGE
    scorer = rs.RougeScorer(["rouge1", "rouge2", "rougeL"], use_stemmer=True)
    r1, r2, rL = [], [], []
    for hyp, ref in zip(hypotheses, references):
        s = scorer.score(ref, hyp)
        r1.append(s["rouge1"].fmeasure * 100)
        r2.append(s["rouge2"].fmeasure * 100)
        rL.append(s["rougeL"].fmeasure * 100)

    return {
        "bleu2":  round(bleu_score, 2),
        "rouge1": round(sum(r1) / len(r1), 2),
        "rouge2": round(sum(r2) / len(r2), 2),
        "rougeL": round(sum(rL) / len(rL), 2),
    }


# ── Main ──────────────────────────────────────────────────────────────────────────────

def print_table(results: dict):
    header = f"{'Model':<30} {'BLEU-2':>8} {'ROUGE-1':>9} {'ROUGE-2':>9} {'ROUGE-L':>9}"
    sep = "-" * len(header)
    print(f"\n{sep}\n{header}\n{sep}")
    for model, scores in results.items():
        marker = " ◄" if model.startswith("ProteinTalk") else ""
        print(f"{model:<30} {scores['bleu2']:>8.2f} {scores['rouge1']:>9.2f} "
              f"{scores['rouge2']:>9.2f} {scores['rougeL']:>9.2f}{marker}")
    print(sep)


def main():
    parser = argparse.ArgumentParser(description="Evaluate ProteinTalk on Mol-Instructions")
    parser.add_argument("--api_url", default="http://localhost:7777",
                        help="ProteinTalk Flask server URL")
    parser.add_argument("--n_samples", type=int, default=50,
                        help="Number of test samples to evaluate (default: 50)")
    parser.add_argument("--out", default="eval_results.json",
                        help="Output JSON file for results")
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--question", default="What is the function of this protein?",
                        help="Question to ask for each protein")
    args = parser.parse_args()

    # Check server is up
    try:
        requests.get(args.api_url, timeout=5)
        print(f"[INFO] ProteinTalk server is reachable at {args.api_url}")
    except Exception:
        print(f"[ERROR] Cannot reach ProteinTalk server at {args.api_url}")
        print("Start the server first:")
        print("  python demo.py --adapter_path ./adapter_weight/adapter_model_and_optimizer_1_400000.pth --use_groq --port 7777 --gpu 0")
        sys.exit(1)

    samples = load_test_samples(args.n_samples, seed=args.seed)

    hypotheses, references, details = [], [], []
    tmp_dir = tempfile.mkdtemp(prefix="proteintalk_eval_")
    print(f"\n[INFO] Temp PDB dir: {tmp_dir}")
    print(f"[INFO] Running evaluation on {len(samples)} samples...\n")

    for i, sample in enumerate(samples):
        sequence = sample["input"].strip()
        reference = sample["output"].strip()
        instruction = sample.get("instruction", args.question).strip()
        # Use a clean, direct question
        question = args.question

        print(f"[{i+1}/{len(samples)}] seq_len={len(sequence)}")

        pdb_path = get_structure(sequence, tmp_dir, i)
        if pdb_path is None:
            print(f"  Skipping (no structure available)")
            continue

        hypothesis = call_proteintalk(args.api_url, pdb_path, question)
        if not hypothesis:
            print(f"  Skipping (empty API response)")
            continue

        print(f"  Ref  : {reference[:120]}...")
        print(f"  Pred : {hypothesis[:120]}...")

        hypotheses.append(hypothesis)
        references.append(reference)
        details.append({
            "sample_idx": i,
            "sequence_length": len(sequence),
            "question": question,
            "reference": reference,
            "hypothesis": hypothesis,
        })

        # Polite delay between API calls
        time.sleep(0.5)

    if not hypotheses:
        print("[ERROR] No valid predictions — check server and network connectivity.")
        sys.exit(1)

    print(f"\n[INFO] Evaluated {len(hypotheses)}/{len(samples)} samples successfully")

    proteintalk_scores = compute_metrics(hypotheses, references)
    proteintalk_scores["n_samples"] = len(hypotheses)

    all_results = {"ProteinTalk (ours)": proteintalk_scores, **PAPER_SCORES}
    print_table(all_results)

    output = {
        "proteintalk_scores": proteintalk_scores,
        "paper_baselines": PAPER_SCORES,
        "config": {
            "n_samples_requested": args.n_samples,
            "n_samples_evaluated": len(hypotheses),
            "question": args.question,
            "api_url": args.api_url,
            "seed": args.seed,
        },
        "details": details,
    }
    with open(args.out, "w") as f:
        json.dump(output, f, indent=2)
    print(f"\n[INFO] Full results saved to: {args.out}")
    print("[INFO] Run update_report.py to add these results to the Word report.")


if __name__ == "__main__":
    main()
