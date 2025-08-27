# Google Menstrual Cycle Calendar 🩸🩸🩸

As a person who menstruates I use a period app to track my cycles. The calendar feature within the app I use is super helpful, but I often forget to check it. 
For ages I’ve thought: *this would be so useful if it could connect to my Google Calendar!*  

Rather than waiting and wondering if the app team will ever create this feature, I decided to build my own solution.  

This is a simple Google Sheets + Apps Script tool that predicts future period cycles and adds them to a dedicated Google Calendar.  

If you’re signed into your Google account, all you need to do is follow [this link](https://docs.google.com/spreadsheets/d/1sMwt2y2igP2e0NYHU-s9dJUS8yLKXtOY2z23H_wxuYU/edit?gid=692630870#gid=692630870?copy) and make a copy of the Google Sheet. 
Then just update it with your own cycle details. The default values are based on my own averages, so they’re a decent place to start if you don’t have all of your own info recorded yet.  

It uses a few values you set (average cycle length, period length, luteal phase, and the start date of your last period) to generate estimates for:  
- ⚡️Pre-bleed week  
- 🌧PMS Day (the day before bleeding starts)  
- 🩸Bleeding days  
- 🥚Ovulation day  
- 🌸Fertile window  

All events are private by default and kept in a separate calendar in your Google account.


## Try it yourself

1. Open the template sheet and **make a copy** into your own Google Drive:  
 [Make a copy](https://docs.google.com/spreadsheets/d/1sMwt2y2igP2e0NYHU-s9dJUS8yLKXtOY2z23H_wxuYU/edit?gid=692630870#gid=692630870?copy) 

2. Go to the **Config** tab and edit the `Values` column to reflect your own cycle:
   - `CalendarName` → e.g. `Cycle`
   - `AverageCycleDays` → e.g. `29`
   - `AveragePeriodDays` → e.g. `4`
   - `LutealDays` → usually `14`
   - `PredictMonthsAhead` → e.g. `3`
   - `AnchorStartDate` → YYYY-MM-DD (first day of your last period)
   - `TitlePrefix` → e.g. `[Cycle]`

3. Reload the sheet. A **Cycle Tools** menu will appear next to *Help* at the top.

4. Use **Cycle Tools → Predict from Settings only** to generate events.

5. In Google Calendar, make sure your `CalendarName` calendar is visible.


## Updating or resetting
- To remove everything the tool created: **Cycle Tools → Delete all created events**  
- To keep it up to date automatically: **Cycle Tools → Install monthly trigger** (rebuilds predictions on the 1st of each month at 06:00).  


## How it works
- The Apps Script (`Code.gs`) reads your config values.  
- It creates a separate Google Calendar (if one doesn’t exist already).  
- Events are tagged with a `TitlePrefix` so they can be safely removed later.  
- All data stays within your own Google account.  


## Screenshots
<img width="1747" height="389" alt="image" src="https://github.com/user-attachments/assets/bbd0cf78-65af-4192-b245-ac91d8edf5e4" />

<img width="1756" height="378" alt="image" src="https://github.com/user-attachments/assets/efc4b371-3bc1-4a48-acb4-100e46f472e7" />

## Feedback & Contributions
Feedback, suggestions, or improvements are always welcome!  
Feel free to [open an issue](../../issues) or message me directly at frankiebdev@gmail.com if you have ideas for how this tool could be better.

## License
[MIT](LICENSE) — free to use, modify, and share.
