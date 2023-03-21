import pandas as pd
import numpy as np
from tkinter import filedialog

csv_file = csv_file =filedialog.askopenfilename()
csv_df = pd.read_csv(csv_file)

print(csv_df)