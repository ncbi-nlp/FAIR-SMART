# Format Converter for PMC‑SM‑BioC

A robust Java tool that converts PubMed Central (PMC) **supplementary materials (SM)** into the interoperable **[BioC](https://bioc.sourceforge.net/)** format for downstream text mining and analysis.
Download the prebuilt  **[JAR](https://ftp.ncbi.nlm.nih.gov/pub/lu/PMCSMBioC/FormatConverter.zip)** .

---

## Overview

The Format Converter ingests supplementary material files (PDF, Word, Excel, CSV/TSV, simple HTML) and emits **BioC XML** (optionally BioC JSON) with consistent structure, provenance, and metadata. It is designed for batch conversion and integration with biomedical NLP pipelines.

## Features

* Multi‑format input: PDF, DOC/DOCX, XLS/XLSX, CSV/TSV, HTML.
* Standards‑compliant BioC output.
* Table linearization with row/column metadata.
* Batch and resumable processing.
* Container‑ready, reproducible runs.

## Supported Formats

**Input → Allowed Output**

* **PDF** → `BioC`, `PubTator`
* **BioC(XML)** → `PubTator`
* **PubTator** → `BioC`
* **PPT / PPTx** → `BioC`, `PubTator`
* **Word (.doc) / Wordx (.docx)** → `BioC`, `PubTator`
* **RTF** → `BioC`
* **Excel (.xls) / Excelx (.xlsx)** → `BioC`, `PubTator`
* **CSV / TSV** → `BioC`
* **TXT (free text)** → `BioC`, `PubTator`
* **IMG (OCR images)** → `BioC`, `PubTator`
* **XML** → `BioC`
* **tar.gz** → **Decompression**

## Installation

```bash
git clone https://github.com/<org>/pmc-sm-bioc-converter.git
cd pmc-sm-bioc-converter
mvn package
```
Or use the prebuilt JAR:

## Usage

```
java -jar FormatConverter.jar [inputfile] [outputfile] [output format:BioC|PubTator] [input format:BioC|PubTator] [fold]
```

**Positional arguments**

* [inputfile] and [outputfile] can be file or folder
* BioC-XML|PubTator|FreeText|PDF|MSWord|MSExcel formats are allowed in [inputfile].
* BioC and PubTator formats are allowed in [outputfile].
* BioC-XML is the default format.
* 
## Output

* BioC collection with documents and passages.
* Metadata includes file type, page/sheet, row/col indices, and checksums.
* Tables are represented as tab‑separated rows.

## Reproducibility

* Fixed dependencies (PDFBox, POI, etc.).
* Containerized execution.
* UTC timestamps and UTF‑8 encoding.

## Citation

If you use **PMC‑SM‑BioC** or this converter, please cite our paper:

> Wei C‑H, Leaman R, Lai P‑T, Comeau D, Tian S, Lu Z. *PMC‑SM‑BioC enables FAIR access to supplementary materials at scale.* PLOS Biology. 2025.

## License

Apache License 2.0. See `LICENSE` for details.








