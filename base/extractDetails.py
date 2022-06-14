def extractDetails(text,count):
    cursor=0
    spaces = 0
    start = 0
    phrase=[]
    while(len(phrase)<count):
        if text[cursor]!=" " and text[cursor]!="\n" and text[cursor]!="/" :
            cursor+=1
            spaces=0
        else:
            if spaces==1:
                spaces=0
                if text[start:cursor].strip()!="":
                    phrase.append(text[start:cursor].strip())
                start=cursor
            else:
                cursor+=1
                spaces+=1

    while(cursor!=len(text)):
        if text[cursor]==" " or text[cursor]=="|" or text[cursor]=="(" or text[cursor]==")" or text[cursor]=="\n":
            if text[start:cursor].strip()!="" and text[start:cursor].strip()!="(" and text[start:cursor].strip()!=")" and text[start:cursor].strip()!="|":
                phrase.append(text[start:cursor].strip().replace("(","").replace(")","").replace("|",""))
            start=cursor
            cursor+=1
        else:
            cursor+=1
    return phrase