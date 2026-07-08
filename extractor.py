#!/usr/bin/env python3
"""PPTX Extractor - Extract content from PowerPoint files"""
import sys
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

def extract_text_from_pptx(path):
    with zipfile.ZipFile(path, 'r') as z:
        slides = sorted([f for f in z.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')])
        for i, slide in enumerate(slides, 1):
            xml_content = z.read(slide)
            root = ET.fromstring(xml_content)
            ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
            texts = []
            for t in root.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
                if t.text:
                    texts.append(t.text)
            print(f"\n{'='*50}")
            print(f"Slide {i}")
            print(f"{'='*50}")
            print(' '.join(texts) if texts else "[Empty slide]")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python extractor.py <path-to-pptx>")
        sys.exit(1)
    extract_text_from_pptx(sys.argv[1])
