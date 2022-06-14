from base import extractDetails


def getDetails(pdf):
    data = {
        "seat_no": [],
        "name": [],
        "prn": [],
        "c1_80": [],
        "c1_20": [],
        "c1_100": [],
        "c1_C": [],
        "c1_G": [],
        "c1_GP": [],
        "c1_C*GP": [],
        "c1_25": [],
        "c1__25": [],
        "c1__C": [],
        "c1__G": [],
        "c1__GP": [],
        "c1__C*GP": [],
        "c2_80": [],
        "c2_20": [],
        "c2_100": [],
        "c2_C": [],
        "c2_G": [],
        "c2_GP": [],
        "c2_C*GP": [],
        "c2_25": [],
        "c2_50": [],
        "c2_75": [],
        "c2__C": [],
        "c2__G": [],
        "c2__GP": [],
        "c2__C*GP": [],
        "c3_80": [],
        "c3_20": [],
        "c3_100": [],
        "c3_C": [],
        "c3_G": [],
        "c3_GP": [],
        "c3_C*GP": [],
        "c3_25": [],
        "c3_50": [],
        "c3_75": [],
        "c3__C": [],
        "c3__G": [],
        "c3__GP": [],
        "c3__C*GP": [],
        "c4_80": [],
        "c4_20": [],
        "c4_100": [],
        "c4_C": [],
        "c4_G": [],
        "c4_GP": [],
        "c4_C*GP": [],
        "c4_25": [],
        "c4__25": [],
        "c4__C": [],
        "c4__G": [],
        "c4__GP": [],
        "c4__C*GP": [],
        "c5_50": [],
        "c5__50": [],
        "c5_100": [],
        "c5_C": [],
        "c5_G": [],
        "c5_GP": [],
        "c5_C*GP": [],
        "c6_50": [],
        "c6__50": [],
        "c6_100": [],
        "c6_C": [],
        "c6_G": [],
        "c6_GP": [],
        "c6_C*GP": [],
        "c7_50": [],
        "c7__50": [],
        "c7_C": [],
        "c7_G": [],
        "c7_GP": [],
        "c7_C*GP": [],
        "Total": [],
        "A-C": [],
        "A-CG": [],
        "GPA": [],
        "pass/fail": [],
        "RPV": [],
        "class": []
    }
    for i in pdf:
        phrase = extractDetails.extractDetails(i, 3)
        data["seat_no"].append(phrase[0])
        data["name"].append(phrase[1].replace("/", ""))
        data["prn"].append(phrase[2])
        # course 1
        data["c1_80"].append(phrase[3])
        data["c1_20"].append(phrase[5])
        data["c1_100"].append(phrase[7])
        data["c1_C"].append(phrase[8])
        data["c1_G"].append(phrase[9])
        data["c1_GP"].append(phrase[10])
        data["c1_C*GP"].append(phrase[11])
        data["c1_25"].append(phrase[12])
        data["c1__25"].append(phrase[14])
        data["c1__C"].append(phrase[15])
        data["c1__G"].append(phrase[16])
        data["c1__GP"].append(phrase[17])
        data["c1__C*GP"].append(phrase[18])
        # course 2
        data["c2_80"].append(phrase[19])
        data["c2_20"].append(phrase[21])
        data["c2_100"].append(phrase[23])
        data["c2_C"].append(phrase[24])
        data["c2_G"].append(phrase[25])
        data["c2_GP"].append(phrase[26])
        data["c2_C*GP"].append(phrase[27])
        data["c2_25"].append(phrase[28])
        data["c2_50"].append(phrase[30])
        data["c2_75"].append(phrase[32])
        data["c2__C"].append(phrase[33])
        data["c2__G"].append(phrase[34])
        data["c2__GP"].append(phrase[35])
        data["c2__C*GP"].append(phrase[36])
        # course 3
        data["c3_80"].append(phrase[38])
        data["c3_20"].append(phrase[40])
        data["c3_100"].append(phrase[42])
        data["c3_C"].append(phrase[43])
        data["c3_G"].append(phrase[44])
        data["c3_GP"].append(phrase[45])
        data["c3_C*GP"].append(phrase[46])
        data["c3_25"].append(phrase[47])
        data["c3_50"].append(phrase[49])
        data["c3_75"].append(phrase[51])
        data["c3__C"].append(phrase[52])
        data["c3__G"].append(phrase[53])
        data["c3__GP"].append(phrase[54])
        data["c3__C*GP"].append(phrase[55])
        # course 4
        data["c4_80"].append(phrase[56])
        data["c4_20"].append(phrase[58])
        data["c4_100"].append(phrase[60])
        data["c4_C"].append(phrase[61])
        data["c4_G"].append(phrase[62])
        data["c4_GP"].append(phrase[63])
        data["c4_C*GP"].append(phrase[64])
        data["c4_25"].append(phrase[65])
        data["c4__25"].append(phrase[67])
        data["c4__C"].append(phrase[68])
        data["c4__G"].append(phrase[69])
        data["c4__GP"].append(phrase[70])
        data["c4__C*GP"].append(phrase[71])
        if phrase[72] != "RPV" and phrase[72] != "ABS":
            # course 5
            data["c5_50"].append(phrase[72])
            data["c5__50"].append(phrase[74])
            data["c5_100"].append(phrase[76])
            data["c5_C"].append(phrase[77])
            data["c5_G"].append(phrase[78])
            data["c5_GP"].append(phrase[79])
            data["c5_C*GP"].append(phrase[80])
            # course 6
            data["c6_50"].append(phrase[81])
            data["c6__50"].append(phrase[83])
            data["c6_100"].append(phrase[85])
            data["c6_C"].append(phrase[86])
            data["c6_G"].append(phrase[87])
            data["c6_GP"].append(phrase[88])
            data["c6_C*GP"].append(phrase[89])
            # course 7
            data["c7_50"].append(phrase[90])
            data["c7__50"].append(phrase[92])
            data["c7_C"].append(phrase[93])
            data["c7_G"].append(phrase[94])
            data["c7_GP"].append(phrase[95])
            data["c7_C*GP"].append(phrase[96])

            data["Total"].append(phrase[100])
            data["A-C"].append(phrase[101])
            data["A-CG"].append(phrase[102])
            data["GPA"].append(phrase[103])

            data["RPV"].append(" ")
        else:
            # course 5
            data["c5_50"].append(phrase[73])
            data["c5__50"].append(phrase[75])
            data["c5_100"].append(phrase[77])
            data["c5_C"].append(phrase[78])
            data["c5_G"].append(phrase[79])
            data["c5_GP"].append(phrase[80])
            data["c5_C*GP"].append(phrase[81])
            # course 6
            data["c6_50"].append(phrase[82])
            data["c6__50"].append(phrase[84])
            data["c6_100"].append(phrase[86])
            data["c6_C"].append(phrase[87])
            data["c6_G"].append(phrase[88])
            data["c6_GP"].append(phrase[89])
            data["c6_C*GP"].append(phrase[90])
            # course 7
            data["c7_50"].append(phrase[91])
            data["c7__50"].append(phrase[93])
            data["c7_C"].append(phrase[94])
            data["c7_G"].append(phrase[95])
            data["c7_GP"].append(phrase[96])
            data["c7_C*GP"].append(phrase[97])

            data["Total"].append(phrase[101])
            data["A-C"].append(phrase[102])
            data["A-CG"].append(phrase[103])
            data["GPA"].append(phrase[104])

            data["RPV"].append(phrase[72])
        data['pass/fail'].append(phrase[37])
        data["class"].append(" ")

    return data
