count = 0
for doc in range(20):
    count +=1
    if int((count / 20 ) * 100 ) % 10 == 0:
        print( count, int((count / 20 ) * 100 ) % 10)