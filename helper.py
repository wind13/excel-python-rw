def find_in_range(value, range):
  for i in range:
    for j in i:
      # print(j.value)
      if j.value == value:
        return j
  return None

log_level = 'debug'

def log(message):
  if log_level == 'debug':
    print(message)
