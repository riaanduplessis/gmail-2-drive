/**
 * The root label to check for sublabels and upload the attachments.
 */
const GMAIL_LABEL: string = 'Finances'

/**
 * The location to store the file.
 *
 * ------------
 * Placeholders
 * ------------
 * $name	The original attachment name
 * $ext	The file extension of the original attachment
 * $domain	The domain part of the sender who sent the attachment
 * $sublabel	The sub label(s) under your configured label where the message was found
 * $y	Year the message was received at
 * $m	Month the message was received at
 * $d	Day the message was received at
 * $h	Hour the message was received at
 * $i	Minute when the message was received at
 * $s	Second when the message was received at
 * $mc	The message number in the thread, starting at 0
 * $ac	The attachment number in the thread, starting at 0
 */
const GDRIVE_FILE: string = 'Finances/Documents/$y/$m/$sublabel/$y$m$d-$domain--$mc-$ac.$ext'

/**
 * The regex to be used for checking the extension of a file name.
 */
const EXTENSION_REGEX: RegExp = /(?:\.([^.]+))?$/

/**
 * Get all the starred threads within our label and process their attachments
 */
function main() {
  let labels = getSubLabels(GMAIL_LABEL)

  for (let i = 0; i < labels.length; i++) {
    let threads = getUnprocessedThreads(labels[i])

    for (let j = 0; j < threads.length; j++) {
      processThread(threads[j], labels[i])
    }
  }
}

/**
 * Returns the Google Drive folder object matching the given path,
 * if the folder doesn't exist it will be created.
 *
 * @param {string} path
 *
 * @return {Folder}
 */ // @ts-ignore Folder
function getOrMakeFolder(path: string): Folder {
  let folder = DriveApp.getRootFolder()
  let names = path.split('/')

  while (names.length) {
    let name = names.shift()
    if (name === '') continue

    let folders = folder.getFoldersByName(name)

    folder = folders.hasNext() ? folders.next() : folder.createFolder(name)
  }

  return folder
}

/**
 * Get the given label and all its sub labels
 *
 * @param {string} name
 *
 * @return {GmailLabel[]}
 */ // @ts-ignore GmailLabel
function getSubLabels(name: string): GmailLabel[] {
  return GmailApp.getUserLabels().filter((label) => {
    return label.getName() === name || // If the label matches the `name` paramter
    label.getName().substr(0, name.length + 1) === name + '/' // If the label is a sub label of the `name` label
  })
}

/**
 * Get all starred threads in the given label
 *
 * @param {GmailLabel} label
 * @return {GmailThread[]}
 */  // @ts-ignore GmailLabel
function getUnprocessedThreads(label: { getThreads: (arg0: number, arg1: number) => any; getName: () => string }): GmailThread[] {
  let from = 0
  let MAX_THREAD_FETCH_PER_LABEL = 50 //maximum is 500
  let threads: string | any[]
  let result = []

  do {
    threads = label.getThreads(from, MAX_THREAD_FETCH_PER_LABEL)
    from += MAX_THREAD_FETCH_PER_LABEL

    for (let i = 0; i < threads.length; i++) {
      if (!threads[i].hasStarredMessages()) continue
      result.push(threads[i])
    }
  } while (threads.length === MAX_THREAD_FETCH_PER_LABEL)

  console.info(result.length + ' threads to process in ' + label.getName())

  return result
}

/**
 * Get the extension of a file based on the constant `EXTENSION_REGEX`.
 *
 * @param  {string} name
 *
 * @return {string}
 */
function getExtension(name: string): string {
  let result = EXTENSION_REGEX.exec(name)

  return result && result[1] ? result[1].toLowerCase() : 'unknown'
}

/**
 * Apply template variables.
 *
 * @param {string} filename with template placeholders
 * @param {info} values to fill in
 *
 * @return {string}
 */
function createFilename(filename: string, info: { [x: string]: any; name?: any; ext?: string; domain?: any; sublabel?: any; y?: string; m?: string; d?: string; h?: string; i?: string; s?: string; mc?: number; ac?: number }) {
  let keys = Object.keys(info)

  keys.sort(function (a, b) {
    return b.length - a.length
  })

  for (let i = 0; i < keys.length; i++) {
    filename = filename.replace(new RegExp('\\$' + keys[i], 'g'), info[keys[i]])
  }

  return filename
}

/**
 * Save a file at the provided path on the User's Google Drive.
 *
 * @param attachment
 * @param path
 *
 * @returns
 */
function saveAttachment(attachment: GoogleAppsScript.Base.BlobSource, path: string) {
  let parts = path.split('/')
  let file = parts.pop()
  path = parts.join('/')

  let folder = getOrMakeFolder(path)

  if (folder.getFilesByName(file).hasNext()) {
    console.warn(path + '/' + file + ' already exists. File not overwritten.')

    return
  }

  folder.createFile(attachment).setName(file)
  console.log(path + '/' + file + ' saved.')
}

/**
 * Process starred messages in a thread and upload the attachments.
 *
 * @param {GmailThread} thread
 * @param {GmailLabel} label where this thread was found
 */
function processThread(thread: { getMessages: () => any }, label: { getName: () => string }) {
  let messages = thread.getMessages()

  for (let j = 0; j < messages.length; j++) {
    let message = messages[j]

    if (!message.isStarred()) continue

    console.info('Processing message from ' + message.getDate())

    let attachments = message.getAttachments()
    for (let i = 0; i < attachments.length; i++) {
      let attachment = attachments[i]
      let info = {
        'name': attachment.getName(),
        'ext': getExtension(attachment.getName()),
        'domain': message.getFrom().split('@')[1].replace(/[^a-zA-Z]+$/, ''), // domain part of email
        'sublabel': label.getName().substr(GMAIL_LABEL.length + 1),
        'y': ('0000' + (message.getDate().getFullYear())).slice(-4),
        'm': ('00' + (message.getDate().getMonth() + 1)).slice(-2),
        'd': ('00' + (message.getDate().getDate())).slice(-2),
        'h': ('00' + (message.getDate().getHours())).slice(-2),
        'i': ('00' + (message.getDate().getMinutes())).slice(-2),
        's': ('00' + (message.getDate().getSeconds())).slice(-2),
        'mc': j,
        'ac': i,
      }

      let file = createFilename(GDRIVE_FILE, info)
      saveAttachment(attachment, file)
    }

    message.unstar()
  }
}
