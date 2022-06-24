# Web Scrapping of Data Comex

The duty of this code is to reduce the time when updating some tables. Before this, everytime new data was available, the update was done manually. That means a lot of time on introducing the parameters, getting the new table and passing the new observations to a personal repository. Therefore, I thought that scrapping the web could drop drastically this time. 

However, web scrapping of this web page, `Data Comex`, is not an easy task. First, no URL was available. Second, as many webs, to get a specific table you must choose some parameters and then the page redirects you to the table. This is a problem because this unables to make petitions to the web since the table that you see, it is not really there.

Thus, the only way to scrap the web was via RSelenium. A package from R (also available in Python) that simulates a navigator (in that case I use Mozella Firefox, but Google Chrome is also available) and from R you can control every action. Search, press, refresh, write. All what you can do in Google Chrome by your own, you can do it in R by writing each action. It is a bit tedious and it may be other optimal solution to this problem. Nevertheless, it appear to be effective and in quite small time, you can get scrapped the table.

The metodology followed is quite easy. I navigate to the home of the web page and I select the parameters desired for each table. It is important to remark that the page has different frames, so you must change repeatedly among them. This task is done with these functions:

- ___findElement(**using** = "", **path** = "")___: you pass it the direction of the element. Open the page source code and search it. You have many options to find the path, `css selector`, `xpath`, `id`, `name`, `class name`, `link text`.
- ___clickElement()___: this function goes after ___findElement()___ and as its name says it reproduces the action of pressing the left button of your mouse.
- ___switchToFrame___: permits you to change the current frame.

Once we reach the table, to extract the information we must take it from the page source. Hence, we use another function:
- ___getPageSource()___

Finally, we have the table in html format, so we transform it to a data frame with the library `rvest`.

For more info, see https://cran.r-project.org/web/packages/RSelenium/RSelenium.pdf.
