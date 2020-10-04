import random
import string

lowercase = string.ascii_lowercase
result_str = ''.join(random.choice(lowercase) for i in range(6))

a = 1