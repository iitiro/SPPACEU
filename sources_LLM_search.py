"""Search text sources with sentence embeddings."""
import argparse
from pathlib import Path

DEFAULT_DATA_DIR = Path(__file__).resolve().parent / "data" / "txt"
DEFAULT_MODEL = "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"


def load_documents(data_dir):
    documents, texts = [], []
    for path in sorted(Path(data_dir).glob("*.txt")):
        text = path.read_text(encoding="utf-8").strip()
        if text:
            documents.append(path.name)
            texts.append(text[:5000])
    if not texts:
        raise ValueError(f"No nonempty UTF-8 .txt sources found in {data_dir}")
    return documents, texts


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data-dir", type=Path, default=DEFAULT_DATA_DIR)
    parser.add_argument("--model", default=DEFAULT_MODEL)
    parser.add_argument("--query", action="append", help="Repeat for multiple queries")
    args = parser.parse_args()
    try:
        documents, texts = load_documents(args.data_dir)
    except ValueError as exc:
        parser.error(str(exc))
    # Load optional dependencies and download model weights only after validating inputs.
    from sentence_transformers import SentenceTransformer
    from sklearn.metrics.pairwise import cosine_similarity

    model = SentenceTransformer(args.model)
    embeddings = model.encode(texts, normalize_embeddings=True)
    queries = args.query or ["публічний аудит", "корупція", "соціальна практика", "протидія корупції", "викриття корупції"]
    for query in queries:
        scores = cosine_similarity(model.encode([query], normalize_embeddings=True), embeddings)[0]
        print(f"\nQUERY: {query}")
        for document, score in sorted(zip(documents, scores), key=lambda item: item[1], reverse=True):
            print(f"{document}: {score:.3f}")


if __name__ == "__main__":
    main()
