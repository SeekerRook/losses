import json
from tabulate import tabulate

prev_years = []
prev_computed = []
prev_results = []
year = 0
IN = []
IN = [17100.04,
41493.56,
26578.88,
22891.24,
106453.52,
74343.41,
-235584.35,
-33947.42,
6829.31,]
for i in IN:

        year+=1
        print("\n\nΕΤΟΣ",year)
        # n = float(input("\n\nΕτήσιο Σύνολο : "))
        n = i
        prev_years.append(n)
        if (n <= 0.0):
            prev_computed.append(n)
        else:
            for i,v in enumerate(prev_computed):
                if i < len(prev_computed) - 5:
                     continue
                if (n*v >0 ):
                    continue
                if (n +v>=0):
                    if i == 0 : print(n+v,n,v)
                    n = n+v
                    prev_computed[i] = 0.0
                else:
                    prev_computed[i] = n+v
                    n = 0.0

                    break
            prev_computed.append(n)
        prev_results.append(sum(i for i in prev_computed[-5:] if i < 0))
        for i,n in enumerate(prev_computed[:-6]): ## remove if not cleaned
                prev_computed[i] = 0
        print("\nΑνάλυση")
        titles= ["Ετος","Πραγματκό Ετήσιο Σύνολο","Σύνολο μετά τις συμψηφίσεις","Ετήσιες Εκπίτπουσες Ζημιές"]

        data=[titles,['']*4]

        for i in range(year):   
            data.append([year,prev_years[i],prev_computed[i],prev_results[i]])
        print(tabulate(data))
        print(f"\n{year}ο Ετος : Ετησιο Συνολο {prev_years[-1]} | Εκπίπτουσες Ζημιές {prev_results[-1]}")

