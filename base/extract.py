from pdfminer.high_level import extract_text
import pandas as pd
from base import mca_sem1, mca_sem2, mca_sem3


def extract_text_from_pdf(pdf):
    pdf = extract_text(open(pdf, 'rb'), caching=True, codec='utf-8')
    pdf = pdf.split(
        '----------------------------------------------------------------------------------------------------------------------------------')
    pdf = [line for line in pdf if line.find("Total Marks obtained") != -1]
    return pdf


def extract(pdf, type):
    pdf = extract_text_from_pdf(pdf)
    if type == "mca_sem1":
        data = mca_sem1.getDetails(pdf=pdf)
    elif type == "mca_sem2":
        data = mca_sem2.getDetails(pdf=pdf)
    elif type == "mca_sem3":
        data = mca_sem3.getDetails(pdf=pdf)
    data = pd.DataFrame(data)
    return data
