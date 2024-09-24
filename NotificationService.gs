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

  maybeSendExpiringItemsNotification('info')
}

/**
 * Warning-level expiring items notification task.
 * @customfunction
 */
function expiringItemsWarningNotificationTask() {
  console.log('Starting expiring items notification task: {"notificationLevel": "warning"}')

  maybeSendExpiringItemsNotification('warning')
}

/**
 * Decides whether or not to send a notification
 * based on the criteria for the passed notification level.
 * @param {string} notificationLevel
 * @customfunction
 */
function maybeSendExpiringItemsNotification(notificationLevel) {
  const medicalSuppliesData = medicalSuppliesSheet
    .getRange(
      1,
      1,
      medicalSuppliesSheet.getLastRow(),
      medicalSuppliesSheet.getLastColumn()
    )
    .getValues()

  const items = getJsonFromArray(medicalSuppliesData)
  const itemsIncludedStatus = filterByStatus(items, excludedStatuses)

  switch (notificationLevel) {
    case 'info': 
      const itemsIncludedStatusInfoDays =
        filterByDaysRemaining(itemsIncludedStatus, '<=', INFO_DAYS)

      if (itemsIncludedStatusInfoDays.length >=1) {
        const itemsIncludedStatusInfoDaysSorted =
          sortByExpirationDate(itemsIncludedStatusInfoDays, 'desc')

        sendExpiringItemsNotification(itemsIncludedStatusInfoDaysSorted, 'info')
      } else {
        console.log(`Expiring items notification task succeeded without sending: {"notificationLevel": "${notificationLevel}", "reason": "no items found matching notification criteria"}`)
      }
      break
    case 'warning':
      const itemsTriggeringWarning =
        filterByDaysRemaining(itemsIncludedStatus, '==', WARNING_DAYS)

      if (itemsTriggeringWarning.length >= 1) {
          const itemsIncludedStatusWarningDays =
            filterByDaysRemaining(itemsIncludedStatus, '<=', WARNING_DAYS)
          const itemsIncludedStatusWarningDaysSorted =
            sortByExpirationDate(itemsIncludedStatusWarningDays, 'desc')
          
          sendExpiringItemsNotification(itemsIncludedStatusWarningDaysSorted, 'warning')
      } else {
        console.log(`Expiring items notification task succeeded without sending: {"notificationLevel": "${notificationLevel}", "reason": "no items found matching notification criteria"}`)
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
