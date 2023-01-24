*** Settings ***
Documentation       Onboarding Producer

Library             RPA.JSON
Library             RPA.Excel.Files
Library             RPA.Robocloud.Items
Library             DateTime
Library             String


*** Tasks ***
producer
    ${config}=    Load JSON from file    config.json
    ${Input}=    Set Variable    ${config}[Input]
    TRY
        ${Data}=    Read input Excel    ${Input}
        Creating dictionaries and uploading as work items    ${Data}
    EXCEPT
        Log    Unable upload workitems pls check the input file
    END


*** Keywords ***
Read input Excel
    [Arguments]    ${Input}
    Open Workbook    ${Input}
    ${Data}=    Read Worksheet As Table    header=True
    RETURN    ${Data}

Creating dictionaries and uploading as work items
    [Arguments]    ${Data}
    FOR    ${row}    IN    @{Data}
        ${Employee ID}=    Set Variable    ${row}[Employee ID]
        ${Email}=    Set Variable    ${row}[Email]
        ${First Name}=    Set Variable    ${row}[First Name]
        ${Last Name}=    Set Variable    ${row}[Last Name]
        ${Date of birth}=    Set Variable    ${row}[Date of birth]
        ${Date of birth}=    Convert To String    ${Date of birth}
        ${Designation}=    Set Variable    ${row}[Designation]
        ${Phone Number}=    Set Variable    ${row}[Phone Number]
        ${Gender}=    Set Variable    ${row}[Gender]

        ${Payload}=    Create Dictionary
        ...    Employee ID=${Employee ID}
        ...    Email=${Email}
        ...    First Name=${First Name}
        ...    Last Name=${Last Name}
        ...    Date of birth=${Date of birth}
        ...    Designation=${Designation}
        ...    Phone Number=${Phone Number}
        ...    Gender=${Gender}

        Create Output Work Item    variables=${Payload}    save=True
    END
    Log    Done
