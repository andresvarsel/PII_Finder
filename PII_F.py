import spacy
import en_core_web_sm

def name_finder(text)

    nlp = spacy.load("en_core_web_sm")

    name_li = []

    with open("data/NameList.txt", "r") as f:
        lines = f.readlines()
        for line in lines:
            name_li.append(line)

    ruler = nlp.add_pipe("entity_ruler", after="ner")

    name_li = [item.strip() for item in name_li]

    patterns = []
    for name in name_li:
        pattern = {"label": "PER", "pattern": name}
        patterns.append(pattern)


    per_li = []
    ruler.add_patterns(patterns)
    doc = nlp(text)

    for ent in doc.ents:
        if ent.label_ == "PERSON":
            per_li.append(ent)

    # print(len(set(per_li)))
    per_li = [str(item) for item in per_li]
    per_li = list(set(per_li))
    per_li = sorted(per_li)
    print(per_li)
    #print(len(set(per_li)))

