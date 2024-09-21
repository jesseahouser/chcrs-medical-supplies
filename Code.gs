const ss = SpreadsheetApp.getActive()
const medicalSuppliesSheet = ss.getSheetByName("Medical Supplies")
const now = new Date()
const MILLIS_PER_DAY = 1000 * 60 * 60 * 24
const excludedStatuses = [
  `Used`,
  `Exchanged`,
  `Destroyed`,
  `Re-purchased`,
]

/**
 * Converts a two dimensional array of spreadsheet data
 * to a JSON array of objects.
 * @param {array} data
 * @return {json}
 * @customfunction
 */
function getJsonFromArray(data) {
  var obj = {}
  var jsonArray = []
  var headers = data[0]
  var cols = headers.length
  var row = []

  for (var i = 1, l = data.length; i < l; i++) {
    obj = {}
    row = data[i]

    for (var col = 0; col < cols; col++) {
      obj[headers[col]] = row[col]
    }
    
    jsonArray.push(obj)
  }

  return jsonArray
}

/**
 * Filters a JSON array of items by comparing against excluded statuses.
 * @param {json} items
 * @param {array} excludedTypes
 * @return {json}
 * @customfunction
 */
function filterByStatus(items, excludedStatuses) {
  const itemsFilteredByStatus = items
    .filter(
      (item) => includesAny(excludedStatuses, item.status.split(', ')) !== true
    )

  return itemsFilteredByStatus
}

/**
 * Returns true if an array includes any value from another array.
 * @param {array} array
 * @param {array} values
 * @return {boolean}
 * @customfunction
 */
function includesAny(array, values) {
  return values.some(v => array.includes(v))
}

/**
 * Filters json array of items by way of:
 *   - the comparison operator
 *   - number of days until expiration
 * @param {json} items
 * @param {string} operator
 * @param {integer} numDays
 * @return {json}
 * @customfunction
 */
function filterByDaysRemaining(items, operator, numDays) {
  const itemsHavingValidExpirationDates = items.filter(
        (item) => isValidDate(item.expiration_date) == true
      )
  switch (operator) {
    case '==':
      return itemsHavingValidExpirationDates.filter(
        (item) => dateDiff(now, item.expiration_date) == numDays
      )
      break
    case '<=':
      return itemsHavingValidExpirationDates.filter(
        (item) => dateDiff(now, item.expiration_date) <= numDays
      )
      break    
  }
}

/**
 * Determines if data passed is a valid date.
 * @param {any} date
 * @return {boolean}
 * @customfunction
 */
function isValidDate(date) {
  if ( Object.prototype.toString.call(date) !== "[object Date]" )
    return false
  return !isNaN(date.getTime())
}

/**
 * Sorts items by expiration date in descending order.
 * @param {json} items
 * @param {string} order
 * @return {json}
 * @customfunction
 */
function sortByExpirationDate(items, order) {
  switch (order) {
    case 'asc' || 'ascending':
      return items.sort((a, b) => b.expiration_date - a.expiration_date)
      break
    case 'desc' || 'descending':
      return items.sort((a, b) => a.expiration_date - b.expiration_date)
      break
  }
}

/**
 * Determines the difference between two dates as follows:
 *   - date2 - date1
 *   - in whole days (rounded up)
 * @param {date} date1
 * @param {date} date2
 * @return {integer}
 * @customfunction
 */
function dateDiff(date1, date2) {
  var time1 = date1.getTime()
  var time2 = date2.getTime()

  return Math.ceil((time2 - time1) / MILLIS_PER_DAY)
}

/**
 * Returns the content from an HTML file.
 * This is used to include a Stylesheet in an HTML file.
 * @param {string} filename
 * @return {string}
 * @customfunction
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

/**
 * Forms the item table body for the HTML email
 * @param {json} items
 * @return {string}
 * @customfunction
 */
function formItemTableBody(items) {
    var itemTableBody = ``

    for (var i = 0; i < items.length; i++) {
      var daysUntilExpiration = dateDiff(now, items[i].expiration_date)
      var formattedExpirationDate = Utilities
        .formatDate(items[i].expiration_date, "EDT", "d MMM yyyy")

      switch (true) {
        case daysUntilExpiration > 14:
          var expirationText = `🟨 ${daysUntilExpiration}`
          break
        case daysUntilExpiration <= 14 && daysUntilExpiration > 0:
          var expirationText = `🟠 ${daysUntilExpiration}`
          break
        case daysUntilExpiration <= 0:
          var expirationText = `‼️ <b>EXPIRED</b>`
          break
      }

      itemTableBody = itemTableBody + `
        <tr>
          <td>${items[i].item}
          <td>${formattedExpirationDate}
          <td>${expirationText}`
    }

    itemTableBody = itemTableBody + `
`

  return itemTableBody
}
