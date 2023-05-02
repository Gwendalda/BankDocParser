git clone https://github.com/Gwendalda/BankDocParser.git
pip install -r BankDocParser/requirements.txt
python BankDocParser/setup.py build
cp BankDocParser/dist/BankDocParser .
rm -rf BankDocParser/