import fitz, sys

doc = fitz.open(r'd:/New folder/Contrast_Enhancement_of_Multiple_Tissues_in_MR_Brain_Images_With_Reversibility.pdf')
with open(r'd:/New folder/paper_text.txt', 'w', encoding='utf-8') as f:
    for i in range(len(doc)):
        f.write(f"\n\n========== PAGE {i+1} ==========\n\n")
        f.write(doc[i].get_text())

print(f"Done. {len(doc)} pages written to paper_text.txt")
