# VBA - Excel

## Overview of Project

### Purpose

The purpose of this project was to refactor the VBA code for improved performance.

## Analysis and Challenges

The original code has nested loops, and the starter code has two for loops. Because the code is nested, it loops through all the data for each ticker. The goal was to modify the code to perform a single loop through the data.

### Analysis of 2017 run

[2017](./resources/VBA_Challenge_2017.png)

### Analysis of 2018 run

[2018](./resources/VBA_Challenge_2018.png)

### Challenges and Difficulties Encountered

The starter code was a bit confusing, and had each loop isolated outside of the nesting in the original. Because of the confusion, I retained the orignal nested loops and included the formatting from the starter code. Even if inefficient, the code delivers the correct desired results. I am interested in seeing a final result without the nesting, to see the approach taken in this example.

## Results

1. Advantages/Disadvantages to the refactored code:
   - I see the issue with the nested loop, but the approach to the starter code was confusing.
   - I believe there would be performance gains from using a single loop, but two were included.

2. Advantages/Disadvantages to the original code:
   - The code works and delivers the correct results, even in inefficient.
   - The inefficient approach of the nested loop makes larger data sets harder to process.
