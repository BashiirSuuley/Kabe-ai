---
title: RAG Document Q&A
emoji: 📚
colorFrom: blue
colorTo: indigo
sdk: gradio
sdk_version: 4.44.0
app_file: app.py
pinned: false
license: mit
---

# RAG Document Q&A System

Upload PDF, DOCX, TXT, or CSV files and ask questions — answers are grounded strictly in your documents.

## Stack
- **Embeddings:** `all-MiniLM-L6-v2` via Sentence-Transformers
- **Vector DB:** FAISS (IndexFlatIP, cosine similarity)
- **LLM:** `llama3-70b-8192` via Groq API
- **UI:** Gradio 4.44

## How to use
1. Get a free API key from [console.groq.com](https://console.groq.com)
2. Paste it in the **Groq API Key** field and click **Authenticate**
3. Upload one or more files and click **Index Documents**
4. Select your **Education Level** and ask questions
