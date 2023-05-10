import pydocparser
import os
import json
import time
from JsonToExcel import jsonParser


def sendFilesToDocParser(files):
    parser = pydocparser.Parser()
    parser.login("f1416f43896a7fa6235c01641b8efa0a40075560")  # api_key
    parserId = "rbc_parser"  # parser_id
    for file in files:
        document_id = parser.upload_file_by_path(file, parserId)
        while True:
            time.sleep(1)
            data = parser.get_one_result(parserId, document_id)
            print(data)
            try:
                if data[0]["id"]:
                    with open(file.replace(".pdf", ".json"), "w") as f:
                        json.dump(data, f)
                    break
            except:
                pass
            # if not data[0]["id"]:
            # print("waiting for the file to be processed...")
            # else:
            # break
        jsonParser(file.replace(".pdf", ".json"))
