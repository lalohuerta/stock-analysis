# stock-analysis

## Project Overview

Steve has asked me to help him run an stock analysis for his parents so they can make the right decision in investment. With all the dataset given I ran a previous analysis finding daily volume and yearly return for only one company DQ nevertheless Steve and his parents want to analyse 11 more companies and compared them.

## Purpose

To make a more efficient code by refactoring so this code will work faster if Steve and his parents want to analyse more companies.

## Results
 
Three different arrays were created to switch the nesting order for all the "for" loops. 

 * tickersVolumes
 * tickersStartingPrices
 * tickersEndingPrices

I created a variable called "tickerIndex" so I can match all the arrays created and iterates through the dataset. With this functions the code ran faster than before and as we added less variables it is using less memory.

### Refactored Code:

[![Refactored-Code-1.png](https://i.postimg.cc/sxm3hqgB/Refactored-Code-1.png)](https://postimg.cc/jnWVV8NK)
[![Refactored-Code-2.png](https://i.postimg.cc/4yfgsPQX/Refactored-Code-2.png)](https://postimg.cc/XGzhgwPP)
[![Refactored-Code-3.png](https://i.postimg.cc/2SsNVr7c/Refactored-Code-3.png)](https://postimg.cc/K1DHWXVB)

* Run-times for Original Code:

[![2017-Original.png](https://i.postimg.cc/MZmsBqqP/2017-Original.png)](https://postimg.cc/LhhBKdqL)
[![2018-Original.png](https://i.postimg.cc/Ss7GcfF2/2018-Original.png)](https://postimg.cc/0MjmsmNP)

* Run-times for Refactored Code:

[![2017-Refactored.png](https://i.postimg.cc/ryDm0F1c/2017-Refactored.png)](https://postimg.cc/ThXGBxNH)
[![2018-Refactored.png](https://i.postimg.cc/25jjzzks/2018-Refactored.png)](https://postimg.cc/18T1B1NB)

This refactored code ran for 2017 = 0.2 & for 2018 = 0.19 seconds faster than the original code. So for this particulary case the refactoring was succesfull.

## Summary 

### General thoughts on refactoring 

### Pros

* It is a very concrete process 
* Could help the original code to be more efficient 
* Could improve legibility of the code 
* Optimization 

### Cons 

* Just reaches the functionality level of the original code
* Could add new bugs to the code if the refactorization is not accurate 

### Refactoring in VBA

One of the biggest advantages of refactoring in VBA is that after the refactoring you will have a more easier to understand or read code and to maintain also that you can use as much as the original code you want to, in the other hand one of the disadvantages of code refactoring in VBA is the syntax, you do not know how much time it may take to complete another processes and due to you do not know all the sintax you maybe end in a situation where you do not know what to do. 











