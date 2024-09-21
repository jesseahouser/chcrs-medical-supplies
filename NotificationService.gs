const INFO_DAYS = 60
const WARNING_DAYS = 14
const emailRecipients = [
  '<INSERT EMAIL ADDRESS 1>',
  '<INSERT EMAIL ADDRESS 2>',
]

/**
 * Info-level expiring items notification task.
 * @customfunction
 */
function expiringItemsInfoNotificationTask() {
  console.log('Starting expiring items notification task: {"notificationLevel": "info"}')
  const medicalSuppliesData = medicalSuppliesSheet
    .getRange(
      1,
      1,
      medicalSuppliesSheet.getLastRow(),
      medicalSuppliesSheet.getLastColumn()
    )
    .getValues()

  const items = getJsonFromArray(medicalSuppliesData)

  maybeSendExpiringItemsNotification(items, 'info')
}

/**
 * Warning-level expiring items notification task.
 * @customfunction
 */
function expiringItemsWarningNotificationTask() {
  console.log('Starting expiring items notification task: {"notificationLevel": "warning"}')
  const medicalSuppliesData = medicalSuppliesSheet
    .getRange(
      1,
      1,
      medicalSuppliesSheet.getLastRow(),
      medicalSuppliesSheet.getLastColumn()
    )
    .getValues()

  const items = getJsonFromArray(medicalSuppliesData)

  maybeSendExpiringItemsNotification(items, 'warning')
}

/**
 * Decides whether or not to send a notification.
 * @param {json} items
 * @param {string} notificationLevel
 * @customfunction
 */
function maybeSendExpiringItemsNotification(items, notificationLevel) {
  switch (notificationLevel) {
    case 'info': 
      var itemsIncludedStatus = filterByStatus(items, excludedStatuses)
      var itemsIncludedStatusInfoDays =
        filterByDaysRemaining(itemsIncludedStatus, '<=', INFO_DAYS)

      if (itemsIncludedStatusInfoDays.length >=1) {
        var itemsIncludedStatusInfoDaysSorted =
          sortByExpirationDate(itemsIncludedStatusInfoDays, 'desc')

        sendExpiringItemsNotification(itemsIncludedStatusInfoDaysSorted, 'info')
      } else {
        console.log(`Expiring items notification task succeeded without sending: {"notificationLevel": "${notificationLevel}"}`)
      }
      break
    case 'warning':
      var itemsIncludedStatus = filterByStatus(items, excludedStatuses)
      var itemsTriggeringWarning =
        filterByDaysRemaining(itemsIncludedStatus, '==', WARNING_DAYS)

      switch (itemsTriggeringWarning.length >= 1) {
        case true:
          var itemsIncludedStatusWarningDays =
            filterByDaysRemaining(itemsIncludedStatus, '<=', WARNING_DAYS)
          var itemsIncludedStatusWarningDaysSorted =
            sortByExpirationDate(itemsIncludedStatusWarningDays, 'desc')
          
          sendExpiringItemsNotification(itemsIncludedStatusWarningDaysSorted, 'warning')
          break
        default:
          console.log(`Expiring items notification task succeeded without sending: {"notificationLevel": "${notificationLevel}"}`)
          break
      }
      break
  }
}

/**
 * Sends a notification email.
 * @param {json} items
 * @param {string} notificationLevel
 * @customfunction
 */
function sendExpiringItemsNotification(items, notificationLevel) {
  const notificationLevelSubject =
    notificationLevel.charAt(0).toUpperCase() + notificationLevel.slice(1)

  switch (notificationLevel) {
    case 'info':
      var emoji = '🗓️'
      var notificationLevelDays = INFO_DAYS.toString()
      break
    case 'warning':
      var emoji = '⚠️'
      var notificationLevelDays = WARNING_DAYS.toString()
      break
  }

  const htmlTemplate = HtmlService.createTemplateFromFile('Email')
  htmlTemplate.notificationLevelDays = notificationLevelDays
  htmlTemplate.itemTableBody = formItemTableBody(items)
  htmlTemplate.sheetUrl = ss.getUrl() + '#gid=' + medicalSuppliesSheet.getSheetId()
  
  MailApp.sendEmail({
    bcc: emailRecipients.toString(),
    name: 'CHCRS Medical Supplies',
    noReply: true,
    subject: `${emoji} [${notificationLevelSubject}] CHCRS Medical Supplies expiring in ${notificationLevelDays} days or less`,
    htmlBody: htmlTemplate.evaluate().getContent()
  })
  console.log(`Expiring items notification sent: {"notificationLevel": "${notificationLevel}", "recipients": [${emailRecipients}]}`)
}
