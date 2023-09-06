const {
  app,
  BrowserWindow,
  shell,
  ipcMain,
  Menu,
  MenuItem,
  clipboard,
} = require("electron");
const settings = require("electron-settings");
const getClientFile = require("./client-injector");
const path = require("path");

let outlookUrl;
let deeplinkUrls;
let mainWindowLoadedPromise
let outlookUrls;
let showWindowFrame;
let $this;

//Setted by cmdLine to initial minimization
const initialMinimization = {
  domReady: false,
};

class MailWindowController {
  constructor() {
    $this = this;
    this.init();
    initialMinimization.domReady = global.cmdLine.indexOf("--minimized") != -1;
  }
  reloadSettings() {
    // Get configurations.
    showWindowFrame =
      settings.getSync("showWindowFrame") === undefined ||
      settings.getSync("showWindowFrame") === true;

    outlookUrl =
      settings.getSync("urlMainWindow") || "https://outlook.office.com/mail";
    deeplinkUrls = settings.getSync("urlsInternal") || [
      "outlook.live.com/mail/deeplink",
      "outlook.office365.com/mail/deeplink",
      "outlook.office.com/mail/deeplink",
      "outlook.office.com/calendar/deeplink",
    ];
    outlookUrls = settings.getSync("urlsExternal") || [
      "outlook.live.com",
      "outlook.office365.com",
      "outlook.office.com",
    ];
    console.log("Loaded settings", {
      outlookUrl: outlookUrl,
      deeplinkUrls: deeplinkUrls,
      outlookUrls: outlookUrls,
    });
  }
  init() {
    this.reloadSettings();

    // Create the browser window.
    this.win = new BrowserWindow({
      x: 100,
      y: 100,
      width: 1400,
      height: 900,
      frame: showWindowFrame,
      autoHideMenuBar: true,

      show: false,
      title: "Prospect Mail",
      icon: path.join(__dirname, "../../assets/outlook_linux_black.png"),
      webPreferences: {
        spellcheck: true,
        nativeWindowOpen: true,
        affinity: "main-window",
        contextIsolation: false,
        nodeIntegration: true,
      },
    });

    // and load the index.html of the app.
    mainWindowLoadedPromise = this.win.loadURL(outlookUrl, { userAgent: "Chrome" });

    // Show window handler
    ipcMain.on("show", (event) => {
      this.show();
    });

    // add right click handler for editor spellcheck
    this.setupContextMenu(this.win);

    // insert styles
    this.win.webContents.on("dom-ready", () => {
      this.win.webContents.insertCSS(getClientFile("main.css"));
      if (!showWindowFrame) {
        this.win.webContents.insertCSS(getClientFile("no-frame.css"));
      }

      this.addUnreadNumberObserver();
      if (!initialMinimization.domReady) {
        this.win.show();
      }
    });

    this.win.webContents.on("did-create-window", (childWindow) => {
      // insert styles
      childWindow.webContents.on("dom-ready", () => {
        childWindow.webContents.insertCSS(getClientFile("main.css"));

        this.setupContextMenu(childWindow);

        let that = this;
        if (!showWindowFrame) {
          let a = childWindow.webContents.insertCSS(
            getClientFile("no-frame.css")
          );
          a.then(() => {
            childWindow.webContents
              .executeJavaScript(getClientFile("child-window.js"))
              .then(() => {
                childWindow.webContents.on("new-window", this.openInBrowser);
                childWindow.show();
              })
              .catch((errJS) => {
                console.log("Error JS Insertion:", errJS);
              });
          }).catch((err) => {
            console.log("Error CSS Insertion:", err);
          });
        }
      });
    });

    // prevent the app quit, hide the window instead.
    this.win.on("close", (e) => {
      //console.log('Log invoked: ' + this.win.isVisible())
      if (this.win.isVisible()) {
        if (
          settings.getSync("hideOnClose") === undefined ||
          settings.getSync("hideOnClose") === true
        ) {
          e.preventDefault();
          this.win.hide();
        }
      }
    });

    // prevent the app minimze, hide the window instead.
    this.win.on("minimize", (e) => {
      if (
        settings.getSync("hideOnMinimize") === undefined ||
        settings.getSync("hideOnMinimize") === true
      ) {
        e.preventDefault();
        this.win.hide();
      }
    });

    // Emitted when the window is closed.
    this.win.on("closed", () => {
      // Dereference the window object, usually you would store windows
      // in an array if your app supports multi windows, this is the time
      // when you should delete the corresponding element.
      this.win = null;
      if (!global.preventAutoCloseApp) {
        app.exit(0); //dont should the app exit is mainWindow is closed?
      }
      global.preventAutoCloseApp = false;
    });

    // Open the new window in external browser
    this.win.webContents.on("new-window", this.openInBrowser);
  }
  addUnreadNumberObserver() {
    this.win.webContents.executeJavaScript(
      getClientFile("unread-number-observer.js")
    );
  }

  toggleWindow() {
    console.log("toggleWindow", {
      isFocused: this.win.isFocused(),
      isVisible: this.win.isVisible(),
    });
    if (/*this.win.isFocused() && */ this.win.isVisible()) {
      this.win.hide();
    } else {
      initialMinimization.domReady = false;
      this.show();
    }
  }
  reloadWindow() {
    initialMinimization.domReady = false;
    this.win.reload();
  }

  openInBrowser(e, url, frameName, disposition, options) {
    console.log("Open in browser: " + url); //frameName,disposition,options)
    if (new RegExp(deeplinkUrls.join("|")).test(url)) {
      // Default action - if the user wants to open mail in a new window - let them.
      //e.preventDefault()
      console.log("Is deeplink");
      options.webPreferences.affinity = "main-window";
    } else if (new RegExp(outlookUrls.join("|")).test(url)) {
      // Open calendar, contacts and tasks in the same window
      e.preventDefault();
      this.loadURL(url);
    } else {
      // Send everything else to the browser
      e.preventDefault();
      shell.openExternal(url);
    }
  }

  show() {
    initialMinimization.domReady = false;
    this.win.show();
    this.win.focus();
  }

  setupContextMenu(tWin) {
    tWin.webContents.on("context-menu", (event, params) => {
      event.preventDefault();
      //console.log('context-menu', params)
      let menu = new Menu();

      if (params.linkURL) {
        menu.append(
          new MenuItem({
            label:
              params.linkURL.length > 50
                ? params.linkURL.substring(0, 50 - 3) + "..."
                : params.linkURL,
            enabled: false,
          })
        );
        menu.append(
          new MenuItem({
            label: "Copy link url",
            enabled: true,
            click: (arg) => {
              clipboard.writeText(params.linkURL, "url");
            },
          })
        );
        menu.append(
          new MenuItem({
            label: "Copy link text",
            enabled: true,
            click: (arg) => {
              clipboard.writeText(params.linkText, "selection");
            },
          })
        );
      }
      //console.log(params)

      for (const flag in params.editFlags) {
        let actionLabel = flag.substring(3); //remove "can"
        if (flag == "canSelectAll") {
          actionLabel = "Select all";
          if (!params.isEditable) {
            continue;
          }
        }
        if (flag == "canUndo" || flag == "canRedo") {
          if (!params.isEditable) {
            continue;
          }
        }
        if (flag == "canEditRichly") {
          continue;
        }
        if (params.editFlags[flag]) {
          menu.append(
            new MenuItem({
              label: actionLabel,
              enabled: true,
              role: flag.substring(3).toLowerCase(),
            })
          );
        }
      }
      if (menu.items.length > 0) {
        menu.popup();
      }
    });
  }
    buildComposeNewMailURL(to, cc, bcc, subject, body) {
        // Outlook Web App (OWA) URL parameters are not documented but the ones below are known to work.
        // Note that OWA treats "cc" and "bcc" as exclusive for some reason; both are ignored if both are present.
        let url = outlookUrl + "/deeplink/compose?popoutv2=1"

        if (to && to.trim().length > 0) {
            url += "&" + "to=" + to
        }
        let sepChar = "?"

        if (cc && cc.trim().length > 0) {
            url += sepChar + "cc=" + cc
            sepChar = "&"
        }

        if (bcc && bcc.trim().length > 0) {
            url += sepChar + "bcc=" + bcc
            sepChar = "&"
        }

        if (subject && subject.trim().length > 0) {
            url += sepChar + "subject=" + subject
            sepChar = "&"
        }

        if (body && body.trim().length > 0) {
            url += sepChar + "body=" + body
            sepChar = "&"
        }

        return url
    }

    doMailToAction(mailToArg) {
        let to
        let cc
        let bcc
        let subject
        let body

        // Remove "mailto:"
        mailToArg = mailToArg.substring(7)

        // Parse mailto based on RFC 6068
        // A "?"" separates the "to" value and others key/value pairs separated by "&"
        // If a "to" value is not specified then the "?" must still be present before key/value pairs
        // "?" is not required if no key/value pairs are present
        let toSeparatorIndex = mailToArg.indexOf("?")
        if (toSeparatorIndex == -1) {
            to = mailToArg
        } else {
            to = mailToArg.substring(0, toSeparatorIndex)
            let kvPairsRaw = mailToArg.substring(toSeparatorIndex + 1)

            let kvPairs = kvPairsRaw.split("&")
            for (const kvPair of kvPairs) {
                let kv = kvPair.split("=", 2)
                if (kv.length != 2) continue
                switch (kv[0].toLowerCase()) {
                    case "cc":
                        cc = kv[1]
                        break
                    case "bcc":
                        bcc = kv[1]
                        break
                    case "subject":
                        subject = kv[1]
                        break
                    case "body":
                        body = kv[1]
                        break
                }
            }
        }

        let newMailURL = this.buildComposeNewMailURL(to, cc, bcc, subject, body)
        mainWindowLoadedPromise.then(() => {
            const newMessageWindow = new BrowserWindow(
                {
                    parent: this.win,
                    x: 200,
                    y: 200,
                    // Dimensions chosen to show all editing controls and no scroll bars as of January 9, 2022
                    width: 1320,
                    height: 750,
                    frame: showWindowFrame,
                    autoHideMenuBar: true,
                    show: false,
                    title: 'Prospect Mail - New Mail',
                    icon: path.join(__dirname, '../../assets/outlook_linux_black.png'),
                    webPreferences: {
                        spellcheck: true,
                        nativeWindowOpen: true
                    }
                })
            newMessageWindow.loadURL(newMailURL)
            newMessageWindow.show()
        })
    }

    getMailToArg(args) {
        for (const arg of args) {
            if (arg.toLowerCase().startsWith("mailto:"))
            {
                return arg
            }
        }

        return undefined
    }

    // Executes command-line actions that can be specified on primary and/or secondary processes
    // Returns true if an action was found and executed, otherwise false
    doCommandLineActions(commandLineArgs) {
        let actionTaken = false

        let mailToArg = this.getMailToArg(commandLineArgs)
        if (mailToArg) {
            console.log('mailto action specified')
            this.doMailToAction(mailToArg)
            actionTaken = true
        }

        return actionTaken
    }

    handleSecondInstance(commandLineArgs) {
        if (!this.doCommandLineActions(commandLineArgs)) this.show()
    }
}

module.exports = MailWindowController;
