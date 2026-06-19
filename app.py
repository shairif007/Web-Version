import os
import io
import zipfile
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import pandas as pd
from pdf2image import convert_from_path
from PIL import Image

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['PROCESSED_FOLDER'] = 'processed'

# Ensure directories exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['PROCESSED_FOLDER'], exist_ok=True)

def optimize_image_size(image, output_path, target_kb):
    target_bytes = target_kb * 1024
    low_quality, high_quality = 5, 95
    best_bytes = None
    while low_quality <= high_quality:
        mid_quality = (low_quality + high_quality) // 2
        img_byte_arr = io.BytesIO()
        image.save(img_byte_arr, format='JPEG', quality=mid_quality)
        size = img_byte_arr.tell()
        if size <= target_bytes:
            best_bytes = img_byte_arr.getvalue()
            low_quality = mid_quality + 1 
        else:
            high_quality = mid_quality - 1 
    if best_bytes:
        with open(output_path, 'wb') as f:
            f.write(best_bytes)
    else:
        image.save(output_path, format='JPEG', quality=5)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    mode = request.form.get('mode')
    files = request.files.getlist('files')
    
    if not files or files[0].filename == '':
        return "No files uploaded", 400

    # Clear previous processing data
    for folder in [app.config['UPLOAD_FOLDER'], app.config['PROCESSED_FOLDER']]:
        for f in os.listdir(folder):
            os.remove(os.path.join(folder, f))

    saved_files = []
    for file in files:
        filename = secure_filename(file.filename)
        path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(path)
        saved_files.append(path)

    output_files = []

    # ==========================================
    # MODULE: EXACT EXCEL MERGE
    # ==========================================
    if mode == "Exact Excel Merge":
        master_data = []
        for file_path in saved_files:
            # Strictly maintain original datatypes to preserve leading zeros
            df = pd.read_excel(file_path, dtype=str)
            df['Source_File'] = os.path.basename(file_path)
            master_data.append(df)
        
        if master_data:
            final_df = pd.concat(master_data, ignore_index=True)
            chk_cols = [c for c in final_df.columns if c != 'Source_File']
            final_df = final_df.drop_duplicates(subset=chk_cols, keep='first')
            
            # Original naming maintained; no timestamps added
            out_path = os.path.join(app.config['PROCESSED_FOLDER'], "Merged_Output.xlsx")
            with pd.ExcelWriter(out_path, engine='xlsxwriter', engine_kwargs={'options': {'strings_to_numbers': False}}) as writer:
                final_df.to_excel(writer, index=False, sheet_name='Sheet1')
            output_files.append(out_path)

    # ==========================================
    # MODULE: SHRINK & CONVERT FILES
    # ==========================================
    elif mode == "Shrink & Convert Files":
        target_kb = int(request.form.get('target_kb', 195))
        for file_path in saved_files:
            try:
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                if file_path.lower().endswith('.pdf'):
                    imgs = convert_from_path(file_path)
                else:
                    imgs = [Image.open(file_path)]
                
                for i, im in enumerate(imgs):
                    im = im.convert('RGB')
                    # Retain original base name without appending system timestamps
                    fname = f"{base_name}.jpg" if len(imgs) == 1 else f"{base_name}_{i}.jpg"
                    out_path = os.path.join(app.config['PROCESSED_FOLDER'], fname)
                    optimize_image_size(im, out_path, target_kb)
                    output_files.append(out_path)
            except Exception as e:
                print(f"Error processing {file_path}: {e}")

    # Zip output files for download (Static naming to avoid system timestamp insertion)
    zip_path = os.path.join(app.config['PROCESSED_FOLDER'], "Processed_Output.zip")
    with zipfile.ZipFile(zip_path, 'w') as zf:
        for f in output_files:
            zf.write(f, os.path.basename(f))
            
    return send_file(zip_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)