#translate Chinese characters into pinyin
import pypinyin
#excel data process library
import xlrd


# 不带声调的(style=pypinyin.NORMAL)
# pinyin without tones

# parameter:  string contains Chinese characters
# return:  pinyin of Chinese characters in the given string
def pinyin(word):
    result = ''
    for i in pypinyin.pinyin(word, style=pypinyin.NORMAL):
        result += ''.join(i)
    return result



def CreateStationFile():

    # the file I would like to store
    with open("test.txt", "w") as f:

        # open the excel file to read data
        xlsx = xlrd.open_workbook('data.xls')

        # this excel file has only one sheet, so we open the first one
        sheet1 = xlsx.sheets()[0]

        # writing json file, the beginning
        f.write("{\n")
        f.write('      "stations" : [\n')

        # since there are 3970 rows of data, the range is from 1 to 3970
        # notice that the data starts at row 2 and ends at row 3970
        for i in range(1, 3970):

            # the 4th column contains the information about the province which the station is located
            province = sheet1.col_values(3)[i]

            # be careful that 陕西 and 山西 have the same pinyin expression
            # so I will let 陕西's pinyin be "shaanxi" and 山西's pinyin be "shanxi"
            # this is also a very common way that how these two provinces are expressed officially in pinyin
            if province == "陕西":
                province = "shaanxi"
            else:
                province = pinyin(province)
            province = province.strip()


            # officially, these five "provinces" are not provinces, but "autonomous region"s
            # e.g.  the expression will become: "xinjiang autonomous region"
            if province == "xinjiang" or province == "xizang" or province == "ningxia" or province == "guangxi" or province == "neimenggu":
                province = province + " autonomous region"
            # officially, these four "provinces" are not provinces, but "centrally-administered municipality"
            # so I just call their names
            elif province == "beijing" or province == "tianjin" or province == "shanghai" or province == "chongqing":
                province = province
            # the rest of the provinces are regular provinces
            # e.g.  the expression will become: "jiangsu province"
            else:
                province = province + " province"

            # the 6th column contains the information about the station name
            station = sheet1.col_values(6)[i]
            station = pinyin(station)



            # writing in this format:
            # {
            #     "province": "xinjiang autonomous region",
            #     "station": "zepu"
            # },
            f.write('          {\n')
            string1 = '              "province" : ' + '"' + province + '"' + ",\n"
            f.write(string1)
            string2 = '              "station" : ' + '"' + station + '"' + "\n"
            f.write(string2)


            # be careful with the edge case
            # we do not need the final ','
            if i != 3969:
                f.write('          },\n')
            else:
                f.write('          }\n')




        # out of loop, write the finishing line
        f.write('      ]\n')
        f.write('  }')

        # since using keyword with open, no need for f.close()




if __name__ == '__main__':
    CreateStationFile()

