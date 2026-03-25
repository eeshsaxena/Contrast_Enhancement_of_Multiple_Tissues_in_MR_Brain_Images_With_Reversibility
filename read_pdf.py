import fitz
doc = fitz.open(r'd:/New folder/Contrast_Enhancement_of_Multiple_Tissues_in_MR_Brain_Images_With_Reversibility.pdf')
pages = [doc[i].get_text() for i in range(len(doc))]
full = '\n===PAGE BREAK===\n'.join(pages)
print(full)
