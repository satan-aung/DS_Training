import json

def save_profile(path, data):
    with open(path,'w',encoding='utf-8') as f:
        json.dump(data,f,indent=2)

def load_profile(path):
    with open(path,'r',encoding='utf-8') as f:
        return json.load(f)
