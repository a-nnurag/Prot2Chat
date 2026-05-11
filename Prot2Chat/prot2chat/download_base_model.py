"""
Downloads Meta-Llama-3-8B-Instruct from HuggingFace into base_model/.

Requirements:
  - A HuggingFace account that has accepted the Llama 3 license at
    https://huggingface.co/meta-llama/Meta-Llama-3-8B-Instruct
  - huggingface-cli login  (run once before this script)

Usage:
  python download_base_model.py
  python download_base_model.py --token hf_xxxx   # pass token directly
"""
import os
import argparse
from huggingface_hub import snapshot_download

REPO_ID = "meta-llama/Meta-Llama-3-8B-Instruct"
DEST = os.path.join(os.path.dirname(os.path.abspath(__file__)), "base_model")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--token", default=None, help="HuggingFace access token")
    args = parser.parse_args()

    print(f"Downloading {REPO_ID} → {DEST}")
    snapshot_download(
        repo_id=REPO_ID,
        local_dir=DEST,
        ignore_patterns=["*.bin", "original/*"],  # skip old .bin shards
        token=args.token,
    )
    print("Done. Base model is ready.")

if __name__ == "__main__":
    main()
