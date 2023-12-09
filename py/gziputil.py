import gzip
from shutil import copyfileobj
import os


for filename in os.listdir("./logs"):
    if filename.endswith(".gz"):
        with gzip.open(os.path.join("logs",filename), "rb") as f_in:
            new_filename = os.path.join("logs",os.path.splitext(filename)[0])
            with open(new_filename,"wb") as f_out:
                print(filename,new_filename)
                copyfileobj(f_in,f_out)