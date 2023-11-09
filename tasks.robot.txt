*** Settings ***
Documentation       Orders robots from RobotSpareBin Industries Inc.
...                 Saves the order HTML receipt as a PDF file.
...                 Saves the screenshot of the ordered robot.
...                 Embeds the screenshot of the robot to the PDF receipt.

Library             RPA.Browser.Selenium
Library             RPA.HTTP
Library             RPA.Tables
Library             RPA.PDF
Library             RPA.Archive
Library             RPA.Dialogs


*** Variables ***
${URL}                  https://robotsparebinindustries.com/#/robot-order
${out_dir}              ${CURDIR}${/}output
${screenshot_dir}       ${out_dir}${/}screenshots
${receipt_dir}          ${out_dir}${/}receipts


*** Tasks ***
Order robots from RobotSpareBin Industries Inc
    Open the robot order website
    ${order_file_url}=    Request order file URL from user
    ${orders}=    Get orders    ${order_file_url}
    Fill the form using the data from the Excel file    ${orders}
    Create ZIP file of all receipts


*** Keywords ***
Open the robot order website
    Open Available Browser    ${URL}
    Click Button    locator=OK

Fill the form using the data from the Excel file
    [Arguments]    ${orders}
    FOR    ${order}    IN    @{orders}
        Fill and submit the form for one order    ${order}
    END

Fill and submit the form for one order
    [Arguments]    ${order}
    Select From List By Value    head    ${order}[Head]
    Click Element    xpath=//input[@value=${order}[Body]]
    ${leg_absolute_xpath}=    Set Variable
    ...    xpath=//div[1]/div[1]/div[1]/div[1]/div[1]/form[1]/div[3]/input[1]
    Input Text    ${leg_absolute_xpath}    ${order}[Legs]
    Input Text    id:address    ${order}[Address]
    Click Button    Preview
    Wait Until Element Is Visible    id:robot-preview-image
    Wait Until Keyword Succeeds    5x    0.5s    Submit next order
    ${pdf}=    Saves the order HTML receipt as a PDF file    ${order}[Order number]
    ${screenshot}=    Saves the screenshot of the ordered robot    ${order}[Order number]
    Embed the robot screenshot to the receipt PDF file    ${screenshot}    ${pdf}
    Click Button    id:order-another
    Click Button    OK

Submit next order
    Click Button    Order
    Wait Until Element Is Visible    id:order-another

Saves the order HTML receipt as a PDF file
    [Arguments]    ${order_number}
    Wait Until Element Is Visible    id:receipt
    Set Local Variable    ${file_path}    ${receipt_dir}${/}receipt_${order_number}.pdf
    ${receipt_results_html}=    Get Element Attribute    id:receipt    outerHTML
    Html To Pdf    ${receipt_results_html}    ${file_path}
    RETURN    ${file_path}

Saves the screenshot of the ordered robot
    [Arguments]    ${order_number}
    Set Local Variable    ${file_path}    ${screenshot_dir}${/}robot_preview_image_${order_number}.png
    Screenshot    id:robot-preview-image    filename=${file_path}
    RETURN    ${file_path}

Embed the robot screenshot to the receipt PDF file
    [Arguments]    ${screenshot}    ${pdf}
    ${image_files}=    Create List    ${screenshot}:align=center
    Add Files To PDF    ${image_files}    ${pdf}    append=True

Create ZIP file of all receipts
    ${zip_file_name}=    Set Variable    ${out_dir}${/}all_receipts.zip
    Archive Folder With Zip    ${receipt_dir}    ${zip_file_name}

Request order file URL from user
    Add heading    Order file URL
    Add text    Please provide the complete URL to the CSV file, which contains all the orders.
    Add text input    url    label=URL for order file
    ${input}=    Run dialog
    RETURN    ${input.url}

Get orders
    [Arguments]    ${order_file_url}
    RPA.HTTP.Download    ${order_file_url}    overwrite=true
    ${orders}=    Read table from CSV    orders.csv
    RETURN    ${orders}
