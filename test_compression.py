import os
from PIL import Image
from compressor import compress_pdf, compress_office_doc, get_file_size_mb, iterative_compress
import zipfile

def create_test_image(path, size=(1000, 1000), color='red'):
    img = Image.new('RGB', size, color=color)
    img.save(path, quality=100)
    return os.path.getsize(path)

def create_test_pdf(path, image_path):
    from pypdf import PdfWriter, PdfReader
    # Extremely simple PDF with one image
    img = Image.open(image_path)
    img.save("temp_page.pdf")
    # Actually, simpler to just use Pillow to save as PDF
    img.save(path, "PDF", resolution=100.0)
    return os.path.getsize(path)

def create_test_docx(path, image_path):
    # Manually create a minimal docx (ZIP with word/media/image1.png)
    temp_dir = "temp_docx_build"
    os.makedirs(os.path.join(temp_dir, "word", "media"), exist_ok=True)
    import shutil
    shutil.copy(image_path, os.path.join(temp_dir, "word", "media", "image1.png"))
    
    # Create [Content_Types].xml and other minimal structures if needed, 
    # but for our compressor, we just need the ZIP structure word/media/
    with zipfile.ZipFile(path, 'w') as zipf:
        zipf.write(os.path.join(temp_dir, "word", "media", "image1.png"), "word/media/image1.png")
    
    shutil.rmtree(temp_dir)
    return os.path.getsize(path)

def test_compression():
    print("Starting tests...")
    
    # Setup test files
    img_path = "test_image.jpg"
    pdf_path = "test_file.pdf"
    docx_path = "test_file.docx"
    output_pdf = "test_file_compressed.pdf"
    output_docx = "test_file_compressed.docx"

    create_test_image(img_path)
    pdf_size = create_test_pdf(pdf_path, img_path)
    docx_size = create_test_docx(docx_path, img_path)

    print(f"Original PDF size: {pdf_size/1024:.2f} KB")
    print(f"Original DOCX size: {docx_size/1024:.2f} KB")

    # Test PDF compression
    iterative_compress(pdf_path, output_pdf, (pdf_size/1024/1024)*0.5, 'pdf')
    comp_pdf_size = os.path.getsize(output_pdf)
    print(f"Compressed PDF size: {comp_pdf_size/1024:.2f} KB")

    # Test DOCX compression
    iterative_compress(docx_path, output_docx, (docx_size/1024/1024)*0.5, 'docx')
    comp_docx_size = os.path.getsize(output_docx)
    print(f"Compressed DOCX size: {comp_docx_size/1024:.2f} KB")

    # Cleanup
    for f in [img_path, pdf_path, docx_path, output_pdf, output_docx, "temp_page.pdf"]:
        try:
            if os.path.exists(f):
                os.remove(f)
        except Exception as e:
            print(f"Warning: Could not remove {f}: {e}")

    if comp_pdf_size < pdf_size and comp_docx_size < docx_size:
        print("Tests PASSED: Compression successful.")
    else:
        print("Tests FAILED: Size did not reduce.")

if __name__ == "__main__":
    test_compression()
