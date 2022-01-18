import numpy as np
data = np.fromfile(r'D:\DataAnalytics\News Filter Project\LE.dat', dtype=float)
print(data.size)
print(data[:200])