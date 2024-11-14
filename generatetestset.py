import xlwings as xw
import pickle
import os
from email import policy
from email.parser import BytesParser
with open("emailset.txt", "r") as emails:
        lines = emails.readlines()
for l in lines:
        as_list = l.split(", ")
len1=len(as_list)

print(len1)
test=[]
hitfile=[]
ws6 = xw.Book("newtrips-annotat.xlsx").sheets['trips&email']
v6b = ws6.range("B1:B111").value

for i in range(1,len1):
        if as_list[i] not in v6b:
                as_list[i].replace("''","")
                test.append(as_list[i])
#print(test[15])


user_input = '/Users/aishwaryavijayan/Documents/emails-podesta'
directory = os.listdir(user_input)
fname=None
n=0

#print(test[1:5].replace("''",""))
for fname in test[15000:17000]:
        fname=fname.lstrip(" \' ")
        fname=fname.rstrip(" \' ")
        print(user_input + os.sep + fname)
        if os.path.isfile(user_input + os.sep + fname):
                print('inside')
                with open(user_input + os.sep + fname, 'rb') as f:
                    name = f.name  # Get file name
                    msg = BytesParser(policy=policy.default).parse(f)
                    print(fname)
          
     
        
                    if(msg.get_body(preferencelist=('plain'))):
                        text = msg.get_body(preferencelist=('plain')).get_content()
                        words = text.split()
                        l=len(words)
                        print(l)
                
                        if(l < 3000):
                    
                            print('getting read')
                            hitfile.insert(n,fname)
                            n=n+1
                        else:
                            print('large, discarded')
        else:
                print('no email')

print(n)
with open('testset15k-17k.pkl', 'wb') as f:
        pickle.dump(hitfile, f)
         
