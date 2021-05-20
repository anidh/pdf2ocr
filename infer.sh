echo "Checking for library support"
sudo apt-get install tesseract-ocr
pip3 install pillow
pip3 install pdf2image
pip3 install pytesseract
pip3 install xlwt
pip3 install pdf2image
python pdf2ocr.py --pdf_name test --excel_name test
echo "removing temp files..."
rm *.png
rm *.txt