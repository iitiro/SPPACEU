from pathlib import Path
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity

DATA_DIR = Path("data/txt")

documents = []
texts = []

for file in DATA_DIR.glob("*.txt"):
    text = file.read_text(encoding="utf-8")

    documents.append(file.name)
    texts.append(text[:3000])  # ограничим размер для теста

print("Loading embedding model...")

model = SentenceTransformer("sentence-transformers/all-MiniLM-L6-v2")

embeddings = model.encode(texts)

print("Embeddings shape:")
print(embeddings.shape)

print("\nSimilarity matrix:\n")

similarities = cosine_similarity(embeddings)

for i, doc1 in enumerate(documents):
    for j, doc2 in enumerate(documents):
        print(f"{doc1} <-> {doc2}: {similarities[i][j]:.3f}")