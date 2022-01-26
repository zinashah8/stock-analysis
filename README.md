# Green Stock Data Analysis 

### Overview:

The purpose of this analysis is to provide our client with an Excel workbook including an easy-to-run VBA macro able to analyze an entire dataset of stocks. This tool will help him determine the best greek stocks via the following analysis. 

The analysis was ran using two VBA scripts: the original script produced through the Module #2 and a refactored version of it. We would then have the opportunity to compare the performances of both scripts, and highlights the pros and cons of refactoring a code.


### Results:

Each stock, except TERP, had a positive return in 2017. 4 stocks had over a 100% return: DQ, ENPH, FSLR and SEDG. The best performing one was DQ with a 199.4% return. 
This was a great performance year for the green stocks.

Most of the green stock returns took a negative turn in 2018. Only ENPH and RUN continued to perform favorably, with 81.9% and 84.0% increases, respectively that year.
ENPH has the highest total daily volume and was the most traded stock that year.

ENPH is globally the best performing green stock over 2017 and 2018.


### Performance Comparison Between Original and Refractored Scripts: 

The analysis was ran with the original VBA script obtained through Module #2 and the refactored version of that script. The idea was to decrease the processing time of the program by only going through all the dataset one time and retrieve all the information.

The original script goes through nested loops over all the data with the variable "j" for each ticker with the variable "i". For 12 tickers, the program goes through the entire dataset 12 times.

The refactored version uses arrays for the results that are filled along going through all the data rows only one time.

* Original 2017 Script Run Time: 0.6337891 sec.
* Original 2018 Script Run Time: 0.5966797 sec.

* Refractored 2017 Script Run Time: 0.1230469 sec. 
* Refractored 2018 Script Run Time: 0.1162109 sec. 

Refactoring the script increase tremendously its performance, as the execution time gets shorter.


### Summary:

Advantages / Disadvantages of Refractoring Code:
* Pros: makes code faster. Preserves a clean and maintainable architecture in evolving code, and reduces bugs.
* Cons: No additional functionality. Costs development time. Quality is highly dependent on previous developer's work.

How these advantages / disadvantages apply to refractoring original VBA script:
* Time to execute function was decreased (more efficient). 
* Code is clearer.
* Code is easier to edit for future analysis.

* Drawbacks: refractored script doesn't allow to analyze a particular set or the whole dataset. Recoding would be necessary to analyze a specific set of the stocks. 