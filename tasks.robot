*** Settings ***
Documentation       Covid Report Extraction.

# libraries for table, browser, excel, email
Library             html_tables.py
Library             RPA.Browser.Selenium          auto_close=${False}
Library             RPA.Tables
Library             RPA.Windows
Library             RPA.Excel.Files
Library             RPA.Email.ImapSmtp    smtp_server=smtp.gmail.com    smtp_port=587
#Task Teardown       Close All Browsers

# variables for the email username, password and recipient
*** Variables  ***
${USERNAME}     aartiak.e5@gmail.com
${PASSWORD}     mzcvowffbjdauwwt
${RECIPIENT}    aartiak.e5@gmail.com


*** Tasks ***
Read HTML table as Table
    # getting html table from the get html table keyword
    ${html_table}=    Get HTML table

    # to convert html table to datatable
    ${table}=    Read Table From Html    ${html_table}

    # setting column name to the given variable
    ${firstcolumn}=    Set Variable    ${0}
    ${Fifthcolumn}=    Set Variable    ${5}
    ${Seventhcolumn}=    Set Variable    ${7}

    # removes the given column from the table
    Pop table column    ${table}   ${firstcolumn}
    Pop table column    ${table}   ${Fifthcolumn}
    Pop table column    ${table}   ${Seventhcolumn}

    # to remove column from 10 to 13 from the table
    ${y}=    Set Variable    ${10}
    WHILE    ${y} < 14
        Log    Executed as long as the condition is True.
        Pop table column    ${table}   ${y}
        ${y}=    Evaluate    ${y} + 1
    END

    # to remove column 15 from the table
    ${x}=    Set Variable    ${15}
    WHILE    ${x} < 22
        Log    Executed as long as the condition is True.
        Pop table column    ${table}   ${x}
        ${x}=    Evaluate    ${x} + 1
    END
    
    # creating the list to rename the table columns
    ${ColumnNames}=    Create List    Location    TotalCases    NewCases    TotalDeath    TotalRecovered    ActiveCases    Serious/Critical    Population        
    Rename Table Columns    ${table}    ${ColumnNames}
   

    # to remove continents row in the table
    ${x}=    Set Variable    ${0}
    ${z}=    Set Variable    ${0}
    WHILE    ${x} < 4
        Log    Executed as long as the condition is True.
        Pop Table Row    ${table}   ${z}
        ${x}=    Evaluate    ${x} + 1
    END

    # to remove south america row in the table
    Pop Table Row    ${table}    1
    
    
    # to remove africa and oceania row in the table
    ${x}=    Set Variable    ${0}
    ${z}=    Set Variable    ${11}
    WHILE    ${x} < 2
        Log    Executed as long as the condition is True.
        Pop Table Row    ${table}   ${z}
        ${x}=    Evaluate    ${x} + 1
    END

    # to remove blank row in the table
    Pop Table Row    ${table}    220

    # to remove total - row in the table
    ${x}=    Set Variable    ${0}
    ${z}=    Set Variable    ${230}
    WHILE    ${x} < 8
        Log    Executed as long as the condition is True.
        Pop Table Row    ${table}   ${z}
        ${x}=    Evaluate    ${x} + 1
    END
    
    # to check the table in log console
    FOR    ${row}    IN    @{table}
        Log To Console    ${row}
    END

    # writing extracted report in csv
    Write table to CSV    ${table}    covidreport.CSV   


Send test email
    # sending email with subject, body and attachments
    Authorize    account=${USERNAME}    password=${PASSWORD}
    Send Message    sender=${USERNAME}
    ...    recipients=${RECIPIENT}
    ...    subject=Covid-19 Case Report Worldwide
    ...    html=True
    ...    body= <html><body><p>Hi Naresh</p></br><p style='padding-left:20px'>I have attached the covid report CSV file for your reference. </p></br><p>Kindly go through it and let me know if you have any questions.</p><p>Thank you.</p></body></html>
    ...    attachments=covidreport.CSV
    

*** Keywords ***
Get HTML table
    #  it opens the given browser
    Open Available Browser
    ...    https://www.worldometers.info/coronavirus/#countries
    ...    maximized=True
    # clicks yesterday button
    Click Link    //a[@href="#nav-yesterday"]
    # get the table for the given locator(selector) and store it in html_table
    ${html_table}=    Get Element Attribute    //table[@id="main_table_countries_yesterday"]    outerHTML
    Log    ${html_table}
    RETURN    ${html_table}
    
    