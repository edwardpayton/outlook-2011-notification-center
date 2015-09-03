# Integrate Outlook 2011 for Mac with Notification Center

Since I'm in the corporate world now, my Mac is one out of probably ten Macs working in a Windows-only environment. That being said, I'm essentially left on my own to figure out hacks and other workarounds to be productive. 

My employer's email, like many others in corporate America, runs off an Exchange Server. Since our version of Exchange isn't supported by Office 365 (not to mention I don't have an Office 365 account) I am left with running Outlook 2011 for Mac rather than the freshly-released Outlook 2016.

The built-in notifications in Outlook 2011 for Mac are well, basically ugly and you cannot force them persist on screen, so I've designed a way to incorporate Outlook 2011 notifications into OSX's Notification Center (with a few caveats).

## Setup
#### 1. Download
First, [download the repo](https://github.com/adamdehaven/outlook-2011-notification-center/archive/master.zip). You will need to save the repository (or at least the `notify.scpt` file) to a location that will not change over time, as you will be referencing it within an Outlook rule.

---

#### 2. Install terminal-notifier
Next, install <a href="https://github.com/julienXX/terminal-notifier" target="_blank">terminal-notifier</a> on your Mac. The *easiest* way to install terminal-notifier is to open up the Terminal app on your Mac and carry out the installation through RubyGems:
   ```
   $ sudo gem install terminal-notifier
   ```  
Otherwise, follow the installation instructions from <a href="https://github.com/julienXX/terminal-notifier" target="_blank">the repository's readme</a>.

---

#### 3. Create the Outlook Rule
Open the **Outlook** application on your Mac. In Outlook's menu, navigate to 
```
Message > Rules > Edit Rules...
```

  1. On the left, select the email account type.  
       *Depending on your email account(s), or if using Exchange, your company's server(s), your selection here will vary depending on your configuration.)*
  2. Click the <kbd>+</kbd> on the bottom right to add a new rule.  
  3. Give the new rule a name (e.g. "Notification Center").
  4. Under the heading "When a new message arrives:" 
     * Add **All Messages** and make sure the selection on the right reads "If all conditions are met".
  5. Under the heading "Do the following:"
     * Choose **Run AppleSCript** and then with the `Script...` file selector, select the `notify.scpt` file downloaded and saved in **Step 1**.
  6. Ensure "Do not apply other rules to messages that meet these conditions" is left **unchecked**, and "Enabled" **is selected**.

When finished, your Outlook rule should look like the screenshot shown below. Once everything looks correct, **click "Ok" to save your rule.**

![Outlook Notification Center Rule screenshot](Outlook-2011-Rule-for-Notification-Center.png?raw=true)

---

#### 4. Options
Customize the type of notifications to display (e.g. 'Banners' or 'Alerts') for your customized Outlook notifications on your Mac
```
System Preferences > Notifications > Microsoft Outlook
```

Clicking on either notification type will open/restore Outlook's main window.

---

#### 5. Caveats
1. Although <a href="https://github.com/julienXX/terminal-notifier" target="_blank">terminal-notifier</a> allows for customizing the notification's icon, the provided method does not currently work in OSX 10.10

2. Clicking on the notification (Banner or Alert) will open/restore Outlook's main window. There is currently no way to tie a script to the specific message that triggered the notification, therefore it's not currently possible (correct me if I'm wrong) to manipulate the specific message outside of adding extra components within the Rule you created in **Step 3**.
