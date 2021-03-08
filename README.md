# VBA Stock Analysis Challenge
  Syed Ahmed 
  March 7, 2020 
  

## Overview 

The purpose of this analysis is to help Steve analyze large data sets containing data about stocks. In this challenge, I am going to be refactoring the original VBA script from module 2 to make the script run faster. 

## Results 

> Upon creating the intial VBA Script from Module 2, I ran some performance tests to see how quickly we can populate results and came up with the following results: 

### 2017 Original 
<img width="1440" alt="2017 before" src="https://user-images.githubusercontent.com/45697471/110263241-11719880-7f84-11eb-994b-e7eec1fd198b.png">

### 2018 Original 
<img width="1132" alt="2018 before" src="https://user-images.githubusercontent.com/45697471/110263250-1afb0080-7f84-11eb-972f-972b35bf5df9.png">

> As we can see from these results, it takes Excel quite a while to run the code to populate the data. 

> After refactoring the code, we were able to improve performance significantly. Here are the results: 

### 2017 Refactored 
<img width="1121" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/45697471/110263308-53024380-7f84-11eb-8fcc-80d7259cfb83.png">

### 2018 Refactored 
<img width="1134" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/45697471/110263322-5bf31500-7f84-11eb-81f8-19aec507f3e4.png">

## Summary 

**What are the advantages or disadvantages of refactoring code?

Refactoring is an important part of the coding process, as it can help increase the efficiency of our scripts as programmers. Making small challenges to the original code can have a significant impact on the overall performance. 

**Advantages**
> Refactoring helps us clean up the code and make sure that it is organized. 
> It can help the code run faster in Excel. 
> Refactoring can also be useful in helping us find bugs and debug them. 

**Disadvantages**
>Refactoring can be a long and tideous process which may involve changing code in multiple locations of a script which can be quite time consuming and costly. 
>It can be risky when developers do not understand the original code. 

**How do these pros and cons apply to refactoring the original VBA script?

>Cleaning up the code and organizing it helped us overcome performance issues with the code, making our script more efficient. We were able to run the code significantly faster over multiple tests. We also were able to debug any errors in the original code that we missed. 
>Initially, when I began the refactoring process it took me quite a while to make it work as there were many parts of the script that needed to be updated in order for the code to run. This was quite tidious at first, but proved to be worth the time. 
