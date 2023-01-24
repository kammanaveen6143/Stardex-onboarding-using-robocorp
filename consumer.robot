*** Settings ***
Documentation       Onboarding consumer

Library             RPA.JSON
Library             RPA.Robocorp.WorkItems
Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             DateTime
Library             RPA.JavaAccessBridge
Library             RPA.Desktop
Library             Collections
Library             RPA.Excel.Files


*** Tasks ***
consumer
    ${config}=    Load JSON from file    config.json
    ${URL}=    Set Variable    ${config}[URL]

    TRY
        Close Workbook
        ${row}=    Set Variable    2
        ${index}=    find index
        open onboarding site    ${URL}
        ${workitem data}=    For Each Input Work Item    Get data from workitem
        ${ststusr}=    Fill the from    ${workitem data}    ${workitem data}    ${index}
        FOR    ${op}    IN    @{ststusr}
            Open Workbook    C:${/}Users${/}kamma.naveen${/}Documents${/}RoboCorp${/}Stardex${/}OnBoarding Form.xlsx
            Set Active Worksheet    Sheet1
            Set Cell Value    ${row}    ${index}    ${op}
            ${row}=    Evaluate    ${row}+1

            Save Workbook
        END
        Close Browser
    EXCEPT
        Log    unable to process the flow due to poor conection
    END


*** Keywords ***
 Get data from workitem
    ${workitem data}=    Get Work Item Payload
    RETURN    ${workitem data}

open onboarding site
    [Arguments]    ${URL}
    TRY
        Open Available Browser
        ...    https://stardex.com.ng/HR-On-boarding-Process-Form/
        ...    maximized=True
        ...    browser_selection=chrome
    EXCEPT
        Log    error
    END

Fill the from
    [Arguments]    ${workitem data}    ${index}    ${onboardstatus}
    ${row}=    Set Variable    2
    ${ststusr}=    Create List
    FOR    ${emp}    IN    @{workitem data}
        ${ID}=    Set Variable    ${emp}[Employee ID]
        ${email}=    Set Variable    ${emp}[Email]
        ${firsname}=    Set Variable    ${emp}[First Name]
        ${Lastname}=    Set Variable    ${emp}[Last Name]
        ${gender}=    Set Variable    ${emp}[Gender]
        ${Dob}=    Set Variable    ${emp}[Date of birth]
        ${Date of birth}=    Convert Date    ${Dob}    %Y-%b-%d
        ${designation}=    Set Variable    ${emp}[Designation]
        ${phone}=    Set Variable    ${emp}[Phone Number]

        IF    ("${ID}" != "None" and "${email}" != "None" and "${firsname}" != "None" and "${Lastname}" != "None" and "${gender}" != "None" and "${Dob}" != "None" and "${designation}" !="None" and "${phone}" != "None")
            Input Text    emp-ID    ${ID}
            # Sleep    1s
            Input Text    emp-email    ${email}
            # Sleep    1s
            Input Text    emp-firstname    ${firsname}
            # Sleep    1s
            Input Text    emp-lastname    ${Lastname}
            #Sleep    1s
            Input Text    emp-dob    ${Date of birth}
            #Sleep    1s
            Input Text    emp-designation    ${designation}
            ##Sleep    1s
            Input Text    phone-number    ${phone}
            #Sleep    1s

            Select From List By Value
            ...    //*[@id="wpcf7-f1759-p1756-o1"]/form/div[2]/div/div[8]/div/p/span/select
            ...    ${gender}
            Select Checkbox    gdpr
            Click Button    //*[@id="wpcf7-f1759-p1756-o1"]/form/div[2]/div/div[10]/p/button
            Sleep    4s
            Wait Until Element Is Visible    //*[@id="wpcf7-f1759-p1756-o1"]/form/div[3]    timeout=20s
            ${Element}=    Run Keyword And Return Status
            ...    Get Webelement
            ...    //*[@id="wpcf7-f1759-p1756-o1"]/form/div[3]
            IF    ${Element} == True
                ${result}=    Capture Element Screenshot
                ...    //*[@id="wpcf7-f1759-p1756-o1"]/form/div[3]
                ...    C:${/}Users${/}kamma.naveen${/}Documents${/}RoboCorp${/}Stardex${/}resultsc${/}${emp}[Employee ID].png
                ${status}=    Get Text    //*[@id="wpcf7-f1759-p1756-o1"]/form/div[3]
                Log    ${status}
                IF    '${status}' == 'One or more fields have an error. Please check and try again.'
                    Log    fields are missing
                    ${onboardstatus}=    Set Variable    Data is missing
                ELSE
                    Log    onboarded sucessfully
                    ${onboardstatus}=    Set Variable    onboarded sucessfully
                END

                Save Workbook
            ELSE
                Log    status not loades in given time

                ${onboardstatus}=    Set Variable    site issue

                CONTINUE
            END
        ELSE
            ${onboardstatus}=    Set Variable    Data is missing
        END
        Append To List    ${ststusr}    ${onboardstatus}
    END

    RETURN    ${ststusr}

find index
    Open Workbook    OnBoarding Form.xlsx
    Set Active Worksheet    Sheet1
    ${table}=    Read Worksheet As Table    header=true

    FOR    ${i}    IN    RANGE    @{table}
        ${counter}=    Set Variable    1
        IF    ${counter} == 1
            ${list}=    Convert To List    ${i}
            Log    ${list}
            ${index}=    Get Index From List    ${list}    Status
            ${counter}=    Evaluate    ${counter}+1
            Log    ${index}
        ELSE
            Log    message
        END
    END
    Log    ${counter}
    ${index}=    Evaluate    ${index}+1
    RETURN    ${index}

set status
    [Arguments]    ${index}    ${row}    ${onboardstatus}
    Open Workbook    C:${/}Users${/}kamma.naveen${/}Documents${/}RoboCorp${/}Stardex${/}OnBoarding Form.xlsx
    Set Active Worksheet    Sheet1
    Set Cell Value    ${row}    ${index}    ${onboardstatus}
    ${row}=    Evaluate    ${row}+1

    Save Workbook
