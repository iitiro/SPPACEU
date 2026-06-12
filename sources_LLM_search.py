from pathlib import Path
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity

DATA_DIR = Path("data/txt")

documents = []
texts = []

for file in DATA_DIR.glob("*.txt"):
    text = file.read_text(encoding="utf-8")

    documents.append(file.name)
    texts.append(text[:5000])

print("Loading model...")

model = SentenceTransformer(
    "sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2"
)

print("Creating embeddings...")

doc_embeddings = model.encode(
    texts,
    normalize_embeddings=True
)

queries = [
    "публічний аудит",
    "корупція",
    "соціальна практика",
    "протидія корупції",
    "викриття корупції"
]

for query in queries:

    print("\n" + "=" * 70)
    print(f"QUERY: {query}")
    print("=" * 70)

    query_embedding = model.encode(
        [query],
        normalize_embeddings=True
    )

    scores = cosine_similarity(
        query_embedding,
        doc_embeddings
    )[0]

    results = list(zip(documents, scores))

    results.sort(
        key=lambda x: x[1],
        reverse=True
    )

    for doc, score in results:
        print(f"{doc}: {score:.3f}")