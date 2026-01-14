# Wee pals compiler

This Python file takes Wee Pals word doc submissions, converts each individually to a PDF file and then compiles all the pdfs into one document.

## Dependencies

Non-standard Python libraries you will need:
* python-docx
* PyPDF2
* docx2pdf
  
## How to use

1. Download the .py file
2. Ensure you have the suitable folder structure
```
project
│ 
└───compiler
│   │   wee_pals_compiler.py
│ 
└───wee_pals_YEAR
    └─── word_docs
    └─── pdfs
```
3. Put all your word docs in the word doc folder
4. Run the .py file from a terminal:
    * navigate to the `compiler` folder
    * execute `python wee_pals_compiler.py [YEAR]` with the year you want to compile

## Convert and compile separately
You can run either function separately, to convert word docs to pdf and to compile pdfs into a single pdf document.
Just add `--func [FUNCTION]` to your execution line, eg `python wee_pals_compiler.py 2025 --func convert`.

## Document file names
The pdf conversion will sort all files alphabetically and order the final document in alphabetic order. To ensure this actually corresponds to participants' names, I usually make sure the word docs follow the convention `wee_pals_[name].docx`.
