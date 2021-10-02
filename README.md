# VBA of Wall Street

## Overview of Project

An analysis of Visual Basic for Applications (VBA) code and its application to a dataset provided by client.

### Purpose

Client approved existing VBA solution work on provided dataset green_stock.xls. Client has requested improvements to the existing macro. Our goal is to repurpose the existing macro code by refactoring to improve speed and efficiency.

## Results

### Updates to Code

The following changes were made to the code:

​	A tickerIndex variable was added and set to zero.

​	Three output arrays were created to hold data.

​		tickerVolumes()

​		tickerStartingPrices()

​		tickerEndingPrices()

The new variables and arrays were created to remove hardcoded values to the code to and to improve the overall efficiency of the macro. These changes also make it more efficient to refactor the code at a later date. It also allows the improvement requested by the client by making the code run more efficiently.

### Outcomes

With the above mentioned changes the outcomes were observed.



For the 2017 dataset, when compared to the green_stocks.xls macro, the changed improved the speed to complete the macro from 0.9023438 (green_stocks.xls) to 0.78125 (VBA_Challenge.xls). An improvement of 13.4199%



For the 2018 dataset, when compared to the green_stocks.xls macro, the changed improved the speed to complete the macro from 0.8867188 (green_stocks.xls) to 0.765625 (VBA_Challenge.xls). An improvement of 13.6564%

These changes should also factor in the efficiency improvements that occurred when applying the formatting changes directly into the macro and not applying the formatting as an additional step after analysis is completed.

## Summary

- What are the advantages or disadvantages of refactoring code?
  - The advantages of refactoring is an obvious improvement in the efficiency of the macro.
  - The disadvantages is the time needed to apply the refactoring to the existing code and debugging and errors that may occur when making changes to existing code.
- How do these pros and cons apply to refactoring the original VBA script?
  - In the long run the process of refactoring is a net win for the user of the macro. It improved the time to analyze the datasets. Although, when making a determination to refactor code an evaluation should be considered if the efficiency outcomes will outweigh the time investment in refactoring the code. 
