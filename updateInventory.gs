{\rtf1\ansi\ansicpg1252\cocoartf1504\cocoasubrtf830
{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww10800\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 function updateInventory() \{\
  var workbook = SpreadsheetApp.getActive()\
  var salesSheet = workbook.getSheetByName('Sales')\
  var salesRange = salesSheet.getDataRange()\
  var salesValues = salesRange.getValues()\
  var purchasesSheet = workbook.getSheetByName('Purchases')\
  var purchasesRange = purchasesSheet.getDataRange()\
  var purchasesValues = purchasesRange.getValues()\
  var inventorySheet = workbook.getSheetByName('Inventory')\
  var inventoryRange = inventorySheet.getDataRange()\
  var inventoryValues = inventoryRange.getValues()\
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
  var inventoryDateCol = 0\
  var inventoryProductCol = 1\
  var inventoryQuantityCol = 2\
  \
  var days = []\
  \
  function sellProduct(dateStr, product, quantity)\{\
   // find date in days\
    var i = 0\
    for(i; i < days.length; i++)\{\
      if(days[i].dateStr == dateStr)\{\
        break\
      \}\
    \}\
    days[i].productSold(product, quantity)\
  \}\
  \
  function buyProduct(dateStr, product, quantity)\{\
   var i =0\
   for(i; i < days.length; i++)\{\
     if(days[i].dateStr == dateStr)\{\
       break\
     \}\
   \}\
    days[i].productPurchased(product, quantity)\
  \}\
  \
  \
  // day object holds date, and two lists\
  // list of products, and corresponding list of quantities\
  // this is super unsafe, don't do it\
  var day = function(date, dateStr)\{\
    this.date = date\
    this.dateStr = dateStr\
    this.productList = []\
    this.productQuantity = []\
    \
    this.productSold = function(productName, quantity)\{\
      if(this.productList.length == 0.0)\{\
        this.productList.push(productName)\
        this.productQuantity.push(-1 * quantity)\
      \}\
      else\{\
        for(var i=0; i<this.productList.length; i++)\{\
          if(this.productList[i] == productName)\{\
            this.productQuantity[i] -= quantity\
            break\
          \}\
          if(i+1 == this.productList.length)\{\
            this.productList.push(productName)\
            this.productQuantity.push(-1 * quantity)\
            break\
          \}\
        \}\
      \}\
    \}\
    this.productPurchased = function(productName, quantity)\{\
      if(this.productList.length == 0.0)\{\
        this.productList.push(productName)\
        this.productQuantity.push(quantity)\
      \}\
      else\{\
        for(var i=0; i<this.productList.length; i++)\{\
          if(this.productList[i] == productName)\{\
            this.productQuantity[i] += quantity\
            break\
          \}\
          if(i+1 == this.productList.length)\{\
            this.productList.push(productName)\
            this.productQuantity.push(quantity)\
            break\
          \}\
        \}\
      \}\
    \}\
  \}\
  \
  function dateSort(a,b)\{\
    return new Date(b.date) - new Date(a.date)\
  \}\
  \
  //Logger.log(salesValues)\
  \
  for (var row=1; row < salesValues.length; row++)\{\
    var date = salesValues[row][salesDateCol]\
    var dateStr = date.toString()\
    var product = salesValues[row][salesProductCol]\
    var quantitySold = salesValues[row][salesQuantityCol]\
\
    // build dates into list\
    if(days.length == 0.0)\{\
      var newDate = new day(date, dateStr)\
      days.push(newDate)\
    \}\
    \
    // see if we already have date in days\
    for(var i=0; i < days.length; i++)\{\
      if (days[i].dateStr == dateStr)\{\
        break\
      \}\
      if(i + 1 == days.length)\{\
        var newDate = new day(date, dateStr)\
        days.push(newDate)\
        break\
      \}\
    \}\
    \
    sellProduct(dateStr, product, quantitySold)\
      \
  \}\
  \
  for (var row = 1; row < purchasesValues.length; row++)\{\
    var date = purchasesValues[row][purchasesDateCol]\
    var dateStr = date.toString()\
    var product = purchasesValues[row][purchasesProductCol]\
    var quantitySold = purchasesValues[row][purchasesQuantityCol]\
\
    // build dates into list\
    if(days.length == 0.0)\{\
      var newDate = new day(date, dateStr)\
      days.push(newDate)\
    \}\
    \
    // see if we already have date in days\
    for(var i=0; i < days.length; i++)\{\
      if (days[i].dateStr == dateStr)\{\
        break\
      \}\
      if(i + 1 == days.length)\{\
        var newDate = new day(date, dateStr)\
        days.push(newDate)\
        break\
      \}\
    \}\
    \
    buyProduct(dateStr, product, quantitySold)\
      \
  \}\
  days.sort(dateSort)\
  days.reverse() // sort function is dumb\
  //Logger.log(days)\
  // write everything out to Inventory Spreadsheet\
  \
  //accumulate totals by day\
  var dateDateList = []\
  var dateProductList = []\
  var dateQuantityList = []\
  \
  var lastProductList = []\
  var lastQuantityList = []\
  \
  var addProductFlag = false\
  \
  var accumulatorQuantityListIndex = 0\
  \
  //build list of all products\
  for(var i=0; i<days.length; i++)\{\
    for(var j=0; j<days[i].productList.length; j++)\{\
      if(dateProductList.length == 0)\{\
        dateProductList.push(days[i].productList[j]) \
      \}\
      else\{\
        addProductFlag = true\
        for(var q=0; q<dateProductList.length; q++)\{\
          if(dateProductList[q] == days[i].productList[j])\{\
           addProductFlag = false\
           break\
          \} \
        \}\
        if(addProductFlag)\{\
         dateProductList.push(days[i].productList[j]) \
        \}\
      \}\
    \}\
  \}\
  \
  //initialize dateQuantityList[0]\
  dateQuantityList[0] = []\
  for(var i=0; i<dateProductList.length; i++)\{\
    dateQuantityList[0].push(0)\
  \}\
  \
  var addQuantityFlag = false\
\
  //fill dateQuantityList\
  for(var i=0; i<days.length; i++)\{\
    dateQuantityList[i+1] = []\
    for(var j=0; j<dateProductList.length; j++)\{\
     addQuantityFlag = true\
     for(var q=0; q<days[i].productList.length; q++)\{\
       if(days[i].productList[q] == dateProductList[j])\{\
        addQuantityFlag = false\
        dateQuantityList[i+1].push(dateQuantityList[i][j]+days[i].productQuantity[q])\
       \}\
     \}\
      if(addQuantityFlag)\{\
       dateQuantityList[i+1].push(dateQuantityList[i][j]) \
      \}\
    \}\
  \}\
  \
  //drop original list of zeroes from dateQuantityList\
  dateQuantityList.splice(0, 1)\
  \
  /*\
  for(var i=0; i<days.length; i++)\{\
    for(var j=0; j<days[i].productList.length; j++)\{\
      dateDateList.push(days[i].date)\
      dateProductList.push(days[i].productList[j])\
      if(lastProductList.length == 0)\{\
        lastProductList.push(days[i].productList[j])\
        lastQuantityList.push(days[i].productQuantity[j])\
        dateQuantityList.push(days[i].productQuantity[j])\
      \}\
      else\{\
        lastProductFlag = false\
        for(var q=0; q<lastProductList.length; q++)\{\
          if(days[i].productList[j] == lastProductList[q])\{\
           lastQuantityList[q] = lastQuantityList[q] + days[i].productQuantity[j]\
           dateQuantityList.push(lastQuantityList[q])\
           lastProductFlag = true\
           break\
          \}\
        \}\
        if(!lastProductFlag)\{ // this logic is wrong\
         lastProductList.push(days[i].productList[j])\
         lastQuantityList.push(days[i].productQuantity[j])\
         dateQuantityList.push(days[i].productQuantity[j])\
        \}\
      \}\
    \}\
  \}\
  */\
  \
  var inventoryRow = 2 //start at row 2\
   /*\
   for(var i=0; i<dateDateList.length; i++)\{\
    inventorySheet.getRange(inventoryRow, inventoryDateCol+1).setValue(dateDateList[i])\
    inventorySheet.getRange(inventoryRow, inventoryProductCol+1).setValue(dateProductList[i])\
    inventorySheet.getRange(inventoryRow, inventoryQuantityCol+1).setValue(dateQuantityList[i])\
    inventoryRow++\
  \}\
  */\
  \
  // write products\
  for(var i=0; i<dateProductList.length; i++)\{\
   inventorySheet.getRange(inventoryRow, inventoryProductCol+i+1).setValue(dateProductList[i]) \
  \}\
  inventoryRow++\
  \
  //write product quantities\
  for(var i=0; i<dateQuantityList.length; i++)\{\
    inventorySheet.getRange(inventoryRow, inventoryDateCol+1).setValue(days[i].date)\
    for(var j=0; j<dateQuantityList[i].length; j++)\{\
     inventorySheet.getRange(inventoryRow, inventoryProductCol+j+1).setValue(dateQuantityList[i][j])\
    \}\
    inventoryRow++\
  \}\
  \
\}}