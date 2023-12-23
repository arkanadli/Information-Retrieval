import os
import docx2txt
import math
import string
import pandas as pd
import tkinter as tk
from PyPDF2 import PdfReader
from collections import Counter
from tkinter import Tk, filedialog, messagebox
from Sastrawi.Stemmer.StemmerFactory import StemmerFactory
from tkinter import ttk
from ttkthemes import ThemedTk  # Import modul untuk tema

# Inisialisasi Stemmer, Stopwords
sastrawi_stemmer = StemmerFactory().create_stemmer()
content_file = docx2txt.process("stopwordlist.docx")
stopwords = set(line.strip() for line in content_file.splitlines() if line.strip())
all_documents = []  # Buat list untuk menyimpan semua dokumen
file_path = ''

def preprocess_text(text, stopwords):
    # Case folding (menyeragamkan kata menjadi lower case semua dan membersihkan punctuation pada teks)
    text = text.lower()
    tokenize = str.maketrans('', '', string.punctuation.replace('-', ''))
    text_without_punct = text.translate(tokenize)
    # Tokenisasi (memisahkan kata dari kalimat)
    words = text_without_punct.split()

    # Filtering (menghapus kata yang berada stop words dan tidak terdapat simbol selain alphanumerik)
    tokens = [word for word in words if word.isalnum() and word not in stopwords]

    # Stemming (menggunakan porter)
    hasil_preprocessing = [sastrawi_stemmer.stem(token) for token in tokens]
    return hasil_preprocessing

# Fungsi untuk membuka file
def open_file(file_path):
    print(file_path)
    os.system(f'start {file_path}')  # Membuka file dengan aplikasi default

# Fungsi untuk menghitung tf
def calculate_tf(word_counts):
    tf = {word: count / sum(word_counts.values()) for word, count in word_counts.items()}
    return tf

# Fungsi untuk menghitung idf
def calculate_idf(documents):
    total_documents = len(documents)
    idf = {word: math.log10(total_documents / sum(1 for doc in documents if word in doc['content'])) for word in set(word for doc in documents for word in doc['content'])}
    return idf

# Fungsi untuk menghitung tf_idf
def calculate_tfidf(tf, idf):
    tfidf = {word: tf[word] * idf[word] for word in tf.keys() & idf.keys()}
    return tfidf

# Fungsi untuk memprosesan utama aplikasi
def process_files_manual():
    global all_documents
    global file_path

    # Cek apakah direktori sudah dipilih
    if not file_path:
        messagebox.showinfo("Warning", "Pilih directory terlebih dahulu")
        return
    # Cek apakah query sudah diisi
    query = query_entry.get()
    if not query:
        messagebox.showinfo("Warning", "Masukkan query terlebih dahulu")
        return

    result_text.delete(1.0, tk.END) # Reset display teks
    all_documents = []  # Reset list dokumen
    print("List of Files in the Directory:")
    for file_name in os.listdir(file_path):
        current_file_path = os.path.join(file_path, file_name)
        print(file_name)

        try:
            file_extension = current_file_path.split('.')[-1].lower()

            if file_extension in ('txt', 'pdf', 'docx'):
                with open(current_file_path, 'rb') as file:
                    raw_content = get_text_from_file(current_file_path, file_extension)
                    content = preprocess_text(raw_content, stopwords)
                    all_documents.append({"content": content, "file_name": file_name, "file_path": current_file_path})
                    word_count = Counter(content)
                    result_text.insert(tk.END, f"Words and Counts in {file_name} (After Preprocessing):\n")
                    for word, count in word_count.items():
                        result_text.insert(tk.END, f"{word}: {count}\n")
                    # Display words and counts using a bar chart

                    result_text.insert(tk.END, f"Total words (After Preprocessing): {len(content)}\n\n")
            else:
                raise ValueError("Unsupported file format")
        except Exception as e:
            print(f"Error processing {file_name}: {e}")

    query = query_entry.get()  # Mengganti penggunaan input() dengan entry GUI
    query_words = preprocess_text(query, stopwords)
    query_word_count = Counter(query_words)
    query_tf = calculate_tf(query_word_count)
    query_idf = calculate_idf(all_documents)
    query_tfidf = calculate_tfidf(query_tf, query_idf)

    relevant_docs_tfidf = []

    for doc in all_documents:
        doc_words = doc['content']
        doc_word_count = Counter(doc_words)
        doc_tf = calculate_tf(doc_word_count)
        print('menghitung tf')
        print(doc_tf)
        doc_idf = calculate_idf(all_documents)
        print('menghitung idf')
        print(doc_idf)
        doc_tfidf = calculate_tfidf(doc_tf, doc_idf)
        print('menghitung tfidf')
        print(doc_tfidf)

        # Calculate cosine similarity between query and document
        dot_product = sum(query_tfidf[word] * doc_tfidf[word] for word in query_tfidf.keys() & doc_tfidf.keys())
        query_norm = math.sqrt(sum(value ** 2 for value in query_tfidf.values()))
        doc_norm = math.sqrt(sum(value ** 2 for value in doc_tfidf.values()))
        print(dot_product)
        # Avoid division by zero
        similarity = dot_product / (query_norm * doc_norm) if query_norm * doc_norm != 0 else 0
        print(similarity)
        print('=======================================')
        relevant_docs_tfidf.append((doc, similarity))

    # Sort relevant documents by similarity in descending order
    relevant_docs_tfidf = sorted(relevant_docs_tfidf, key=lambda x: x[1], reverse=True)

    # Menampilkan dokumen dengan bobot similiritas
    result_text.insert(tk.END, "Urutan Dokumen dengan similiritas tertinggi:\n\n")
    for i, (doc, similarity) in enumerate(relevant_docs_tfidf):
        result_text.insert(tk.END, f"Rank: {i + 1}\n")
        result_text.insert(tk.END, f"Nilai Similaritas: {similarity:.4f}\n")
        result_text.insert(tk.END, f"File Name: {doc['file_name']}\n")
        result_text.insert(tk.END, f"Path: {doc['file_path']}\n")
        result_text.insert(tk.END, "-" * 150 + "\n")

        # Tambahkan tombol "Open File" untuk membuka file terkait
        open_button = ttk.Button(result_text, text="Open File", command=lambda path=doc['file_path']: open_file(path))
        result_text.window_create(tk.END, window=open_button)
        result_text.insert(tk.END, "\n\n")

def browse_directory():
    global file_path
    file_path = filedialog.askdirectory(title="Select Directory")

    if file_path:
        file_path_label.config(text=f"Selected Directory: {file_path}")

# Fungsi untuk mendapatkan teks dari file
def get_text_from_file(file_path, file_extension):
    if file_extension == 'txt':
        with open(file_path, 'r', encoding='utf-8') as file:
            return file.read()
    elif file_extension == 'pdf':
        with open(file_path, 'rb') as file:
            pdf_reader = PdfReader(file)
            content = ''
            for page_num in range(len(pdf_reader.pages)):
                content += pdf_reader.pages[page_num].extract_text()
            return content
    elif file_extension == 'docx':
        return docx2txt.process(file_path)

# Membuat windows aplikasi menggunakan tkinter
root = Tk()
root.withdraw()
root = ThemedTk(theme="radiance")
# root = ThemedTk(theme="scidmint")
# root = ThemedTk(theme="breeze")
root.title("Document Search App")
frame = ttk.Frame(root, padding="10")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
file_path_label = ttk.Label(frame, text="Pilih Direktori Kumpulan Dokumen: ", font=("Helvetica", 12))
file_path_label.grid(row=1, column=0, columnspan=2, pady=5)
browse_button = ttk.Button(frame, text="Cari Direktori", command=browse_directory, width=30)
browse_button.grid(row=2, column=0, columnspan=2, pady=15,padx=15)

query_entry_label = ttk.Label(frame, text="Masukkan Query:   ", font=("Helvetica", 12))
query_entry_label.grid(row=3, column=0, pady=30, sticky="e")

query_entry = ttk.Entry(frame, width=40, font=("Helvetica", 12))
query_entry.grid(row=3, column=1, pady=5, sticky="w")

search_button = ttk.Button(frame, text="Cari Dokumen", command=process_files_manual,width=30)
search_button.grid(row=4, column=0, columnspan=2, pady=20)

label = ttk.Label(frame, text="Hasil:   ", font=("Helvetica", 12))
label.grid(row=5, column=0, pady=0, sticky='w')

result_text = tk.Text(frame, wrap="word", width=120, height=20, padx=15,pady=15, font=("Helvetica", 12))
result_text.grid(row=6, column=0, columnspan=2, pady=10)

# Menjalankan loop utama GUI
root.mainloop()
