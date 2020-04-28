def doWork(filename = None):
    print("Starting process. Please be alert.")
    import win32com.client
    shell = win32com.client.Dispatch("WScript.Shell")
    import time
    import pyautogui

    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys

    # https://sites.google.com/a/chromium.org/chromedriver/home
    # Download ChromeDriver and extract. Then enter the path here.
    PATH_TO_CHROME = r"PATH\TO\CHROMEDRIVER\chromedriver.exe"

    # The link to your Google Sheet
    GOOGLE_SHEET_LINK = "https://docs.google.com/spreadsheets/d/1bsbcPkyy5DkZjBCSpW5dJYdcAy6ncRtBwPts0gaTGXU/"
    # https://docs.google.com/spreadsheets/d/1bsbcPkyy5DkZjBCSpW5dJYdcAy6ncRtBwPts0gaTGXU/

    NUM_COLUMNS_TOSELECT = 2
    NUM_ROWS_TOSELECT = 3

    browser = webdriver.Chrome(PATH_TO_CHROME)
    res = browser.get(GOOGLE_SHEET_LINK)


    userPrompt = "N"
    while(userPrompt != "y" and userPrompt != "Y"):
        userPrompt = input("Did you sign into Google, and open the Spreadsheet? (Y or y to continue) ")

    print (res)


    ## Checks in case the file is View-Only. Exits the loop only when it changes.
    ## Great for countdowns. 
    # isWrite = False
    # while (isWrite == False):
    #     print ("still view only")
    #     time.sleep(0.5)
    #     try:
    #         if (browser.find_elements_by_xpath("//*[contains(text(), '\"View Only\"')]") == False):
    #             isWrite = True
    #     except:
    #         isWrite == False



    print ("PUT YOUR MOUSE ON THE STARTING CELL NOW!!! STARTING IN...")
    
    for i in range (0, 3):
        print(3-i, "...", sep="")
        time.sleep(1)

    print ("GO!")

    pyautogui.click()

    for i in range (0, NUM_COLUMNS_TOSELECT-1):
        shell.SendKeys("+{RIGHT}")

    for i in range (0, NUM_ROWS_TOSELECT-1):
        shell.SendKeys("+{DOWN}")

    pyautogui.click(button='right')

    time.sleep(0.5)

    for i in range (0, 15):
        shell.SendKeys("{DOWN}")

    shell.SendKeys("{ENTER}")

    time.sleep(1)

    for i in range (0, 6):
        shell.SendKeys("{TAB}", 0)

    shell.SendKeys("{ENTER}")

    popModalHasOpened = False
    while (popModalHasOpened == False):
        try:
            loadingBar = browser.find_element_by_class_name("waffle-ritz-protection-acl-loading-spinner")
            if (loadingBar.is_displayed()==True):
            # waffle-ritz-protection-acl-loading-spinner
                while (popModalHasOpened == False):
                    if (loadingBar.is_displayed()==False):
                        print("WWWWWWWWWWWWWWWOOOOOOOOOOOOOOOOOOOOHHHHHHHHHHHHHHHOOOOOOOOOOOOOOOOO")
                        popModalHasOpened = True
            else:
                print("Isn't display:none'd yet")
        except:
            print("Not open")
            popModalHasOpened = False

    time.sleep(2)

    for i in range (0, 4):
        shell.SendKeys("{TAB}", 0)
        time.sleep(0.3)
    time.sleep(0.5)
    shell.SendKeys("{ENTER}")
    time.sleep(0.5)

    shell.SendKeys("{UP}")
    # time.sleep(1)

    shell.SendKeys("{ENTER}")
    time.sleep(0.5)

    shell.SendKeys("{TAB}", 0)
    # time.sleep(1)

    shell.SendKeys("{ENTER}")


    # Thinking time
    time.sleep(3)



doWork()