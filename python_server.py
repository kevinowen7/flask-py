from flask import Flask, request, jsonify

import docx
import os
from io import BytesIO
from datetime import datetime
from PIL import Image
import pytesseract

import firebase_admin
from firebase_admin import credentials
from firebase_admin import storage

import io

app = Flask(__name__)

pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe'

cred = credentials.Certificate('serviceAccountKey.json')
firebase_admin.initialize_app(cred, {
    'storageBucket': 'project-ta-4e64c.appspot.com'
})
bucket = storage.bucket()

@app.route('/')
def hello_world():
    return 'Hello, World!'

@app.route('/convert-image', methods=['POST'])
def convert_image():
    if 'file' not in request.files:
        return jsonify(
            success=0,
            message="File not found",
            data=None
        )
    
    image = request.files['file']

    # upload origin file to firebase
    filename_original = datetime.now().strftime('%Y%m%d%H%M%S')+".jpg"

    f = io.BytesIO()
    pil_img = Image.open(image)
    pil_img.save(f, format='JPEG')
    pil_img.close()
    
    blob_original = bucket.blob("origin/"+filename_original)
    blob_original.upload_from_string(f.getvalue())
    blob_original.make_public()
    
    # OCR
    text_parsed = pytesseract.image_to_string(Image.open(image))

    # Create a document
    doc = docx.Document()

    # Add a paragraph to the document
    p = doc.add_paragraph()
    p.add_run(text_parsed)

    # save document
    filename = datetime.now().strftime('%Y%m%d%H%M%S')+".docx"
    doc.save(filename)
    
    # upload to firebase
    blob = bucket.blob("result/"+filename)
    blob.upload_from_filename(filename)
    blob.make_public()
    os.unlink(filename)

    resultData = {}
    resultData["result_file"] = blob.public_url
    resultData["original_file"] = blob_original.public_url
    
    # END
    return jsonify(
        success=1,
        message="Success",
        data=resultData
    )

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=8080)
