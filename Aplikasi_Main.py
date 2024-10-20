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

# Inisialisasi Stemmer, Stopwords , dan juga Basewords
sastrawi_stemmer = StemmerFactory().create_stemmer()
content_file = docx2txt.process("stopwordlist.docx")
stopwords = set(line.strip() for line in content_file.splitlines() if line.strip())
content_file2 = docx2txt.process("list_kata.docx")
basewords = set(line.strip() for line in content_file2.splitlines() if line.strip())
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

    # Menentukan apakah kata tersebut sudah kata dasar atau belum
    # Stemming (menggunakan sastrawi), tambahkan pengecekan sebelum stemming
    sastrawi_stemmer = StemmerFactory().create_stemmer()
    hasil_preprocessing = []
    for token in tokens:
        # Cek apakah kata sudah merupakan kata dasar sebelum stemming
        if token in basewords:
            hasil_preprocessing.append(token)
        else:
            hasil_preprocessing.append(sastrawi_stemmer.stem(token))

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
    relevant_docs_tfidf = [] # Reset list tf-idf
    result_text.insert(tk.END, f"-------Hasil Preprocessing Teks------\n")
    result_text.insert(tk.END, f"Case Folding -> Tokenisasi -> Filtering -> Stemming \n")
    result_text.insert(tk.END, f"\n")
    for file_name in os.listdir(file_path):
        current_file_path = os.path.join(file_path, file_name)

        try:
            file_extension = current_file_path.split('.')[-1].lower()

            if file_extension in ('txt', 'pdf', 'docx'):
                with open(current_file_path, 'rb') as file:
                    raw_content = get_text_from_file(current_file_path, file_extension)
                    content = preprocess_text(raw_content, stopwords)
                    all_documents.append({"content": content, "file_name": file_name, "file_path": current_file_path})
                    word_count = Counter(content)
                    result_text.insert(tk.END, f"Kata dasar dan kemunculan di file : {file_name}\n")
                    for word, count in word_count.items():
                        result_text.insert(tk.END, f"{word}: {count}\n")
                    # Display words and counts using a bar chart

                    result_text.insert(tk.END, f"Total Kata (After Preprocessing): {len(content)}\n\n")
            else:
                raise ValueError("Unsupported file format")
        except Exception as e:
            print(f"Error processing {file_name}: {e}")

    # Proses Pengolahan kata dari query yang diinputkan user
    query = query_entry.get()  # Mendapatkan query dari entry GUI
    query_words = preprocess_text(query, stopwords)  # Melakukan preprocessing teks pada query
    query_word_count = Counter(query_words)  # Menghitung frekuensi kata pada query
    query_tf = calculate_tf(query_word_count)  # Menghitung Term Frequency (TF) pada query
    query_idf = calculate_idf(all_documents)  # Menghitung Inverse Document Frequency (IDF) untuk semua dokumen
    query_tfidf = calculate_tfidf(query_tf, query_idf)  # Menghitung TF-IDF untuk query

    for doc in all_documents:
        doc_words = doc['content']
        doc_word_count = Counter(doc_words) # Menghitung frekuensi kata pada doc
        doc_tf = calculate_tf(doc_word_count) # Menghitung Term Frequency (TF) pada doc
        doc_idf = calculate_idf(all_documents) # Menghitung Inverse Document Frequency (IDF) untuk semua dokumen
        doc_tfidf = calculate_tfidf(doc_tf, doc_idf) # Menghitung TF-IDF untuk doc

        # Calculate cosine similarity between query and document
        dot_product = sum(query_tfidf[word] * doc_tfidf[word] for word in query_tfidf.keys() & doc_tfidf.keys())
        query_norm = math.sqrt(sum(value ** 2 for value in query_tfidf.values()))
        doc_norm = math.sqrt(sum(value ** 2 for value in doc_tfidf.values()))

        # Avoid division by zero
        similarity = dot_product / (query_norm * doc_norm) if query_norm * doc_norm != 0 else 0

        relevant_docs_tfidf.append((doc, similarity))

    # Sort relevant documents by similarity in descending order
    relevant_docs_tfidf = sorted(relevant_docs_tfidf, key=lambda x: x[1], reverse=True)

    result_text.insert(tk.END, f"-------Hasil Temu Balik Informasi------\n")
    result_text.insert(tk.END, f"TF-IDF -> Cosine Similitary  \n")
    result_text.insert(tk.END, f"\n")
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
# root = ThemedTk(theme="radiance")
# root = ThemedTk(theme="scidmint")
root = ThemedTk(theme="breeze")
root.title("Document Search App")
frame = ttk.Frame(root, padding="10")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
file_path_label = ttk.Label(frame, text="Pilih Direktori Kumpulan Dokumen: ", font=("Helvetica", 12))
file_path_label.grid(row=1, column=0, columnspan=2, pady=5)
browse_button = ttk.Button(frame, text="Cari Direktori", command=browse_directory, width=30)
browse_button.grid(row=2, column=0, columnspan=2, pady=15,padx=15)

# Kolom query
query_entry_label = ttk.Label(frame, text="Masukkan Query:   ", font=("Helvetica", 12))
query_entry_label.grid(row=3, column=0, pady=30, sticky="e")

query_entry = ttk.Entry(frame, width=40, font=("Helvetica", 12))
query_entry.grid(row=3, column=1, pady=5, sticky="w")

# Button Cari Dokumen
search_button = ttk.Button(frame, text="Cari Dokumen", command=process_files_manual,width=30)
search_button.grid(row=4, column=0, columnspan=2, pady=20)

# Panel Display Hasil
label = ttk.Label(frame, text="Hasil:   ", font=("Helvetica", 12))
label.grid(row=5, column=0, pady=0, sticky='w')

result_text = tk.Text(frame, wrap="word", width=120, height=20, padx=15,pady=15, font=("Helvetica", 12))
result_text.grid(row=6, column=0, columnspan=2, pady=10)

# Menjalankan loop utama GUI
root.mainloop()
