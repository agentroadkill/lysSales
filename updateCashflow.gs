{\rtf1\ansi\ansicpg1252\cocoartf1504\cocoasubrtf830
{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww10800\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 function updateCashflow() \{\
  \
  var workbook = SpreadsheetApp.getActive()\
  var salesSheet = workbook.getSheetByName('Sales')\
  var salesRange = salesSheet.getDataRange()\
  var salesValues = salesRange.getValues()\
  var purchasesSheet = workbook.getSheetByName('Purchases')\
  var purchasesRange = purchasesSheet.getDataRange()\
  var purchasesValues = purchasesRange.getValues()\
  var cashSheet = workbook.getSheetByName('Cashflow')\
  var cashRange = cashSheet.getDataRange()\
  var cashValues = cashRange.getValues()\
  var productSheet = workbook.getSheetByName('Product')\
  var productRange = productSheet.getDataRange()\
  var productValues = productRange.getValues()\
  \
  \
  \
  var salesDateCol = 0\
  var salesProductCol = 2\
  var salesQuantityCol = 3\
  var salesAdjPriceCol = 4\
  \
  var purchasesDateCol = 0\
  var purchasesProductCol = 1\
  var purchasesQuantityCol = 2\
  var purchasesAdjPriceCol = 3\
  \
  var cashDateCol = 0\
  var cashIncomeCol = 1\
  var cashExpenseCol = 2\
  var cashNetCol = 3\
  \
  var productNameCol = 0\
  var productCostCol = 1\
  var productPriceCol = 2\
  var productMarginCol = 3\
  \
  var days = []\
  \
  var day = function(date, dateStr)\{\
   \
    this.date = date\
    this.dateStr = dateStr\
    \
    this.income = 0\
    this.expense = 0\
    \
    this.net = function ()\{\
      this.income - this.expense\
    \}\
    \
    this.sell = function(ammount)\{\
     this.income = this.income + ammount \
    \}\
    \
    this.buy = function(ammount)\{\
     this.expense = this.expense + ammount \
    \}\
    \
  \}\
  \
  var product = function(name, price, cost)\{\
 \
    this.name = name\
    this.price = price\
    this.cost = cost\
    \
  \}\
  \
  function dateSort(a,b)\{\
    return new Date(b.date) - new Date(a.date)\
  \}\
  \
  //build list of all products\
  var productList = []\
  productList.push(0)\
  //start at 1\
  for(var row=1; row<productValues.length; row++)\{\
    var name = productValues[row][productNameCol]\
    //look for duplicates\
    for(var i=0; i<productList.length; i++)\{\
      if(productList[i].Name == name)\{\
       var ui = SpreadsheetApp.getUi()\
       ui.alert("Error!","Duplicate product values found")\
       exit(1)\
      \}\
    \}\
    \
    //add products to list\
    var price = productValues[row][productPriceCol]\
    var cost = productValues[row][productCostCol]\
    var newProduct = new product(name,price,cost)\
    productList.push(newProduct)\
  \}\
  productList.splice(0, 1)\
  \
  //start building days\
  var price = 0\
  days[0] = 0\
  \
  for(var row=1; row<salesValues.length; row++)\{\
    var date = salesValues[row][salesDateCol]\
    var dateStr = date.toString()\
    var product  = salesValues[row][salesProductCol]\
    var adjPrice = salesValues[row][salesAdjPriceCol]\
    var needNewDate = true\
    \
    for(var i=0; i<days.length; i++)\{\
      if(days[i].dateStr == dateStr)\{\
        if(adjPrice.toString() == '')\{ //because apparently (0.0 =='') -> true\
          for(var q=0; q<productList.length; q++)\{\
            if (product == productList[q].name)\{\
             price = productList[q].price\
             break\
            \}\
          \}\
        \}\
        else\{\
          price = adjPrice\
        \}\
        Logger.log("Selling "+product+" on "+date+" for $"+price)\
        days[i].sell(price)\
        needNewDate = false\
      \}\
    \}\
    \
    if(needNewDate)\{\
      var newDay = new day(date, dateStr)\
      days.push(newDay)\
      \
      if(adjPrice == '')\{\
        for(var q=0; q<productList.length; q++)\{\
          if (product == productList[q].name)\{\
            price = productList[q].price\
            break\
          \}\
        \}\
      \}\
      else\{\
        price = adjPrice\
      \}\
      Logger.log("Selling "+product+" on "+date+" for $"+price)\
      days[days.length-1].sell(price)\
    \}\
  \}\
  \
  for (var row=1; row<purchasesValues.length; row++)\{\
    var date = purchasesValues[row][purchasesDateCol]\
    var dateStr = date.toString()\
    var product  = purchasesValues[row][purchasesProductCol]\
    var adjPrice = purchasesValues[row][purchasesAdjPriceCol]\
    var needNewDate = true\
    \
    //look for date\
    for (var i=0; i<days.length; i++)\{\
      needNewDate = true\
      if (dateStr == days[i].dateStr)\{\
        needNewDate = false\
        //determine price\
        if (adjPrice == '')\{\
          for(var j=0; j<productList.length; j++)\{\
             if (product == productList[i].name)\{\
             price = productList[i].price\
             break\
            \}\
          \}\
        \}\
        else\{\
         price = adjPrice \
        \}\
        \
        //buy product\
        days[i].buy(price)\
        \
      \}\
    \}\
    \
    if(needNewDate)\{\
      var newDay = new day(date, dateStr)\
      days.push(newDay)\
      \
      if(adjPrice == '')\{\
        for(var q=0; q<productList.length; q++)\{\
          if (product == productList[q].name)\{\
            price = productList[i].price\
            break\
          \}\
        \}\
      \}\
      else\{\
        price = adjPrice\
      \}\
      days[days.length-1].buy(price)\
    \}\
    \
  \}\
    \
  //drop first dummy date\
  days.splice(0, 1)\
  \
  //trigger net update for each day\
  for(var i=0; i<days.length; i++)\{\
   days[i].net() \
  \}\
  \
  \
  //sort days\
  days.sort(dateSort)\
  days.reverse()\
  //Logger.log(productList)\
  //Logger.log(days)\
  \
  //write data out to spreadshet\
  for(var i=0; i<days.length; i++)\{\
    cashSheet.getRange(i+2, cashDateCol+1).setValue(days[i].date)\
    cashSheet.getRange(i+2, cashIncomeCol+1).setValue(days[i].income)\
    cashSheet.getRange(i+2, cashExpenseCol+1).setValue(days[i].expense)\
    cashSheet.getRange(i+2, cashNetCol+1).setValue(days[i].income - days[i].expense) //net isn't working right now\
  \}\
  //write out totals\
  cashSheet.getRange(i+2, cashDateCol+1).setValue("Total")\
  cashSheet.getRange(i+2, cashDateCol+1).setFontWeight("bold")\
  cashSheet.getRange(i+2, cashIncomeCol+1).setBackground("green")\
  cashSheet.getRange(i+2, cashIncomeCol+1).setFormulaR1C1("=SUM(R["+-i+"]C["+(cashIncomeCol-1)+"]:R["+-1+"]C["+(cashIncomeCol-1)+"])")\
  cashSheet.getRange(i+2, cashExpenseCol+1).setBackground("red")\
  cashSheet.getRange(i+2, cashExpenseCol+1).setFormulaR1C1("=SUM(R["+-i+"]C["+(cashExpenseCol-2)+"]:R["+-1+"]C["+(cashExpenseCol-2)+"])")\
  cashSheet.getRange(i+2, cashNetCol+1).setFormulaR1C1("=SUM(R["+-i+"]C["+(cashNetCol-3)+"]:R["+-1+"]C["+(cashNetCol-3)+"])")\
  \
\}\
\
}