# Excel Data into PPTX 
A Python Project to convert data from excel into a PowerPoint report  

# About

During my internship at State Street, as a data analyst, I was able to implement a system to turn our project manager’s data/schedule on Excel into a PowerPoint report. This PowerPoint was a progress report for the entire team to see the schedule of each project and their respective detail. Before I created this program project managers were required to dedicated 2-3 hours during their work week to build their indvidual report slide. While I am not able to show the project that I worked on during my internship, I instead spent some time over winter break to showcase what I did during my time in the summer. Furthermore, instead of a PowerPoint on project management I decided to make it about the National Parks. 

# Requirements 

- Python 3 
- Excel 
- PowerPoint 

# Installation 
- python-pptx
- wikipedia
- urllib.request (beautiful soup)


# Description 
This program takes data from a [Wikipedia page](https://en.wikipedia.org/wiki/List_of_national_parks_of_the_United_States) listing the national parks. Using Excel's Get Data feature we can get data from the wikipedia page. However, this feature does not download photo's or photo url's so I used the Beautiful Soup, a webscraper library, to get a list of urls from the Wikipedia page as a single string. Both sets of data required a bit of cleaning using pandas. Afterwards I had to use python pptx to set up a slide template of where I would like all my information to go for each of the state parks. Then I would just loop through the list of national parks to create a slide deck. Since the slide template was a bit plain I decided to use a theme to add some quick color. Shapes can be implimented into the powerpoint if more detail is desired. 

# Files 
- Excel file with the wikipedia data 
- Powerpoint with manually added theme 
- excel to pptx.py python code 
