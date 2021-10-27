import json

#read file json
file = open("data/Properties.json")
obj = json.load(file)

print(obj["AddressEmail"])