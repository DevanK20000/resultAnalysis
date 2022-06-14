from base.extractDetails import extractDetails


def getDetails(pdf):
    data = {
        "elective1": [],
        "elective2": [],
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
        "c4_80": [],
        "c4_20": [],
        "c4_100": [],
        "c4_C": [],
        "c4_G": [],
        "c4_GP": [],
        "c4_C*GP": [],
        "c4_25": [],
        "c4_50": [],
        "c4_75": [],
        "c4__C": [],
        "c4__G": [],
        "c4__GP": [],
        "c4__C*GP": [],
        "c5_80": [],
        "c5_20": [],
        "c5_100": [],
        "c5_C": [],
        "c5_G": [],
        "c5_GP": [],
        "c5_C*GP": [],
        "c5_25": [],
        "c5__25": [],
        "c5__C": [],
        "c5__G": [],
        "c5__GP": [],
        "c5__C*GP": [],
        "c6_50": [],
        "c6_C": [],
        "c6_G": [],
        "c6_GP": [],
        "c6_C*GP": [],
        "c7_50": [],
        "c7__50": [],
        "c7_100": [],
        "c7_C": [],
        "c7_G": [],
        "c7_GP": [],
        "c7_C*GP": [],
        "c8_25": [],
        "c8_50": [],
        "c8_75": [],
        "c8_C": [],
        "c8_G": [],
        "c8_GP": [],
        "c8_C*GP": [],
        "c9_25": [],
        "c9_50": [],
        "c9_75": [],
        "c9_C": [],
        "c9_G": [],
        "c9_GP": [],
        "c9_C*GP": [],
        "c10_50": [],
        "c10_C": [],
        "c10_G": [],
        "c10_GP": [],
        "c10_C*GP": [],
        "Total": [],
        "A-C": [],
        "A-CG": [],
        "GPA": [],
        "pass/fail": [],
        "RPV": [],
        "class": [],
    }
    for i in pdf:
        phrase = extractDetails(i, 6)
        data["elective1"].append(phrase[0])
        data["elective2"].append(phrase[1])
        data["seat_no"].append(phrase[2])
        data["name"].append(phrase[3].replace("/", ""))
        data["prn"].append(phrase[4])
        # course 1
        data["c1_80"].append(phrase[6])
        data["c1_20"].append(phrase[8])
        data["c1_100"].append(phrase[10])
        data["c1_C"].append(phrase[11])
        data["c1_G"].append(phrase[12])
        data["c1_GP"].append(phrase[13])
        data["c1_C*GP"].append(phrase[14])
        data["c1_25"].append(phrase[15])
        data["c1__25"].append(phrase[16])
        data["c1__C"].append(phrase[17])
        data["c1__G"].append(phrase[18])
        data["c1__GP"].append(phrase[19])
        data["c1__C*GP"].append(phrase[20])
        # course 2
        data["c2_80"].append(phrase[21])
        data["c2_20"].append(phrase[23])
        data["c2_100"].append(phrase[25])
        data["c2_C"].append(phrase[26])
        data["c2_G"].append(phrase[27])
        data["c2_GP"].append(phrase[28])
        data["c2_C*GP"].append(phrase[29])
        data["c2_25"].append(phrase[30])
        data["c2_50"].append(phrase[32])
        data["c2_75"].append(phrase[34])
        data["c2__C"].append(phrase[35])
        data["c2__G"].append(phrase[36])
        data["c2__GP"].append(phrase[37])
        data["c2__C*GP"].append(phrase[38])
        # course 3
        data["c3_80"].append(phrase[40])
        data["c3_20"].append(phrase[42])
        data["c3_100"].append(phrase[44])
        data["c3_C"].append(phrase[45])
        data["c3_G"].append(phrase[46])
        data["c3_GP"].append(phrase[47])
        data["c3_C*GP"].append(phrase[48])
        # course 4
        data["c4_80"].append(phrase[49])
        data["c4_20"].append(phrase[51])
        data["c4_100"].append(phrase[53])
        data["c4_C"].append(phrase[54])
        data["c4_G"].append(phrase[55])
        data["c4_GP"].append(phrase[56])
        data["c4_C*GP"].append(phrase[57])
        data["c4_25"].append(phrase[58])
        data["c4_50"].append(phrase[60])
        data["c4_75"].append(phrase[62])
        data["c4__C"].append(phrase[63])
        data["c4__G"].append(phrase[64])
        data["c4__GP"].append(phrase[65])
        data["c4__C*GP"].append(phrase[66])
        if phrase[67] != "RPV" and phrase[67] != "ABS":
            # course 5
            data["c5_80"].append(phrase[67])
            data["c5_20"].append(phrase[69])
            data["c5_100"].append(phrase[71])
            data["c5_C"].append(phrase[72])
            data["c5_G"].append(phrase[73])
            data["c5_GP"].append(phrase[74])
            data["c5_C*GP"].append(phrase[75])
            data["c5_25"].append(phrase[76])
            data["c5__25"].append(phrase[77])
            data["c5__C"].append(phrase[78])
            data["c5__G"].append(phrase[79])
            data["c5__GP"].append(phrase[80])
            data["c5__C*GP"].append(phrase[81])
            # course 6
            data["c6_50"].append(phrase[82])
            data["c6_C"].append(phrase[83])
            data["c6_G"].append(phrase[84])
            data["c6_GP"].append(phrase[85])
            data["c6_C*GP"].append(phrase[86])
            # course 7
            data["c7_50"].append(phrase[87])
            data["c7__50"].append(phrase[89])
            data["c7_100"].append(phrase[91])
            data["c7_C"].append(phrase[92])
            data["c7_G"].append(phrase[93])
            data["c7_GP"].append(phrase[94])
            data["c7_C*GP"].append(phrase[95])
            # course 8
            data["c8_25"].append(phrase[96])
            data["c8_50"].append(phrase[98])
            data["c8_75"].append(phrase[100])
            data["c8_C"].append(phrase[101])
            data["c8_G"].append(phrase[102])
            data["c8_GP"].append(phrase[103])
            data["c8_C*GP"].append(phrase[104])
            # course 9
            data["c9_25"].append(phrase[105])
            data["c9_50"].append(phrase[107])
            data["c9_75"].append(phrase[109])
            data["c9_C"].append(phrase[110])
            data["c9_G"].append(phrase[111])
            data["c9_GP"].append(phrase[112])
            data["c9_C*GP"].append(phrase[113])

            # course 10
            data["c10_50"].append(phrase[114])
            data["c10_C"].append(phrase[115])
            data["c10_G"].append(phrase[116])
            data["c10_GP"].append(phrase[117])
            data["c10_C*GP"].append(phrase[118])

            data["Total"].append(phrase[122])
            data["A-C"].append(phrase[123])
            data["A-CG"].append(phrase[124])
            data["GPA"].append(phrase[125])

            data["RPV"].append(" ")
        else:
            # course 5
            data["c5_80"].append(phrase[68])
            data["c5_20"].append(phrase[70])
            data["c5_100"].append(phrase[72])
            data["c5_C"].append(phrase[73])
            data["c5_G"].append(phrase[74])
            data["c5_GP"].append(phrase[75])
            data["c5_C*GP"].append(phrase[76])
            data["c5_25"].append(phrase[77])
            data["c5__25"].append(phrase[78])
            data["c5__C"].append(phrase[79])
            data["c5__G"].append(phrase[80])
            data["c5__GP"].append(phrase[81])
            data["c5__C*GP"].append(phrase[82])
            # course 6
            data["c6_50"].append(phrase[83])
            data["c6_C"].append(phrase[84])
            data["c6_G"].append(phrase[85])
            data["c6_GP"].append(phrase[86])
            data["c6_C*GP"].append(phrase[87])
            # course 7
            data["c7_50"].append(phrase[88])
            data["c7__50"].append(phrase[90])
            data["c7_100"].append(phrase[92])
            data["c7_C"].append(phrase[93])
            data["c7_G"].append(phrase[94])
            data["c7_GP"].append(phrase[95])
            data["c7_C*GP"].append(phrase[96])
            # course 8
            data["c8_25"].append(phrase[97])
            data["c8_50"].append(phrase[99])
            data["c8_75"].append(phrase[101])
            data["c8_C"].append(phrase[102])
            data["c8_G"].append(phrase[103])
            data["c8_GP"].append(phrase[104])
            data["c8_C*GP"].append(phrase[105])
            # course 9
            data["c9_25"].append(phrase[106])
            data["c9_50"].append(phrase[108])
            data["c9_75"].append(phrase[110])
            data["c9_C"].append(phrase[111])
            data["c9_G"].append(phrase[112])
            data["c9_GP"].append(phrase[113])
            data["c9_C*GP"].append(phrase[114])

            # course 10
            data["c10_50"].append(phrase[115])
            data["c10_C"].append(phrase[116])
            data["c10_G"].append(phrase[117])
            data["c10_GP"].append(phrase[118])
            data["c10_C*GP"].append(phrase[119])

            data["Total"].append(phrase[123])
            data["A-C"].append(phrase[124])
            data["A-CG"].append(phrase[125])
            data["GPA"].append(phrase[126])

            data["RPV"].append(phrase[67])
        data['pass/fail'].append(phrase[39])
        data["class"].append(" ")
    return data
