from num2words import num2words

def number_to_words(n):
    return num2words(n, lang='en_IN').replace(",", "").title() + " Rupees Only"
