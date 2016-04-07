
# Sim's convert2txt script

### Installation

Open a Terminal and change to the directory you want to install convert2txt:

For example: 

```
cd ~/git/
```

Then run these commands:

```
git clone https://github.com/SimKennedy/convert2txt.git
cd convert2txt
virtualenv venv
source venv/bin/activate
pip install -r requirements.txt
```

### Using convert2txt

When you want to run the script:

```
cd ~/git/convert2txt
source venv/bin/activate
python convert2txt.py path/to/folder/or/file/to/convert
```



### Using ExtractMSG

```
cd ~/git/convert2txt
source venv/bin/activate
python extract_msg.py [path to dir]
```