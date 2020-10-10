import os
import random

file_id = '_'+str(random.randint(999, 9999))
filename = 'files/asd.pdf'
filename = os.path.splitext(filename)[0] + file_id + os.path.splitext(filename)[1]
pass