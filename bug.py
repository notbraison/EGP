from playwright.sync_api import sync_playwright

captured = []

def log_response(response):
    try:
        ctype = response.headers.get("content-type") or ""
        if "json" in ctype or "app" in response.url.lower():
            captured.append(response)
    except Exception:
        pass

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()

    page.route("**/deskpro-messenger/**", lambda route: route.abort())
    page.goto("https://egpkenya.go.ke/public-app", wait_until="domcontentloaded", timeout=60000)
    page.wait_for_selector("table tbody tr", timeout=30000)

    print("Page loaded. Now clicking first row link...")
    page.on("response", log_response)

    link = page.query_selector_all("table tbody tr")[0].query_selector_all("td")[3].query_selector("a")
    link.click()

    page.wait_for_timeout(4000)

    print("URL after click:", page.url)
    print("Captured", len(captured), "responses fired during/after click")
    for r in captured:
        print("URL:", r.url)
        try:
            body = r.text()
            print(body[:2000])
        except Exception as e:
            print("body unreadable:", e)
        print("---")

    input("Press Enter to close browser...")
    browser.close()
    
    
    ##get-app-details/28506 leaks real PII. That response includes a real user's login ID, email, phone number, and a bcrypt password hash for the officer who created the plan. You don't need that endpoint at all for your use case (procurement entity/budget data)
    ## Captured 3 responses fired during/after click
##URL: https://egpkenya.go.ke/api/app/get-app-details/28506
##{"appDetailId":28506,"financialYearId":5,"appRefNo":"TENP/864/APP/2026-27/3","typeOfApp":3,"createdOn":"2026-07-09T13:33:17","createdByUserId":7760,"cbUserDesDeptHistId":62403,"updatedOn":"2026-07-09T16:37:03","updatedByUserId":7760,"ubUserDesDeptHistId":62403,"xstatus":1,"activationRemarks":null,"activatedOn":null,"activatedByUserid":null,"abUserDesDeptHistId":null,"appStage":null,"isConsolidated":true,"reviewerUserId":null,"peDepartmentId":864,"adminFinancialYear":{"financialYearId":5,"financialYear":"2026-27","startDate":"2026-07-01","endDate":"2027-06-30","createdOn":"2025-07-01T16:11:24","createdByUserId":1,"cbUserDesDeptHistId":1,"isCurrentFinancialYear":true,"xstatus":1,"activatedon":null,"activatedByUserId":null,"abUserDesDeptHistId":null,"hibernateLazyInitializer":{}},"cbUserRegister":{"id":7760,"loginId":"25695030","password":"$2a$10$4ezjOEp3ticw1dgZpsKu.u66qtllFnwN0AFAKy98l67/T1IlfoQMm","userType":null,"emailId":"charlesrutto789@gmail.com","secondaryEmailId":null,"isdCode":"+254","mobileNo":"724010789","secondaryMobileNo":null,"namePrefix":null,"firstName":"CHARLES","middleName":"KIPROP","lastName":"RUTTOH","gender":{"genderId":1,"gender":"Male","hibernateLazyInitializer":{}},"citizenshipTypeId":null,"citizenshiPrefNo":"25695030","employeeId":"2021N00001","registatus":null,"isFirstLogin":false,"failedLoginAttempt":0,"isAccountLocked":false,"accountLockedOn":null,"lastLoginIpAddress":"192.168.203.211","lastLoginDate":"2026-09-08T10:38:10","isPasswordChanged":true,"passwordExpiryDate":"2026-11-24T12:03:45","isTncAccepted":true,"createDon":"2025-07-02T12:03:45","createdByUserId":5383,"cbUserDesDeptHistId":1308,"updateDon":"2026-02-10T13:23:10","updatedByUserId":282,"ubUserDesDeptHistId":24797,"activateDon":"2025-07-02T12:03:45","activatedByUserId":5383,"abUserDesDeptHistId":1308,"activationRemarks":null,"xstatus":1,"hibernateLazyInitializer":{}},"cbUserPedesigndeptHistory":{"id":62403,"userId":7760,"peDesignationId":1396,"userRegiHistoryId":197368,"createdOn
---
URL: https://egpkenya.go.ke/api/app/get-appdetail-summary/28506
##{"reportdata":[{"financialyear":"2026-27","ministryname":"MINISTRY OF EDUCATION","pedepartmentid":864,"apptype":3,"appstatus":"Approved","createdon":"2026-07-09","amended":"No","apprefno":"TENP/864/APP/2026-27/3","createdby":"CHARLES KIPROP RUTTOH - Procurement Manager - Head of Procurement (HoP)","directorate":null,"procentityname":"ELDORET NATIONAL POLYTECHNIC","depname":"ELDORET NATIONAL POLYTECHNIC","isamended":false,"typeofapp":"APP/HoP LEVEL APP"}]}

##URL: https://egpkenya.go.ke/api/app/view-app-summary
##{"status":1,"respData":{"totalcount":55,"reportdata":[{"unspscsegment":"44000000","description":"Office Equipment & Accessories & Supplies","updatedTotalCost":21613200.00,"totalCount":55,"totalCost":21613200.00,"typeofapp":3},{"unspscsegment":"56000000","description":"Furniture & Furnishings","updatedTotalCost":6605220.00,"totalCount":55,"totalCost":6605220.00,"typeofapp":3},{"unspscsegment":"39000000","description":"Electrical systems & Lighting & components & accessories & supplies","updatedTotalCost":3362400.00,"totalCount":55,"totalCost":3362400.00,"typeofapp":3},{"unspscsegment":"43000000","description":"Information Technology Broadcasting & Telecommunications","updatedTotalCost":38793180.00,"totalCount":55,"totalCost":38793180.00,"typeofapp":3},{"unspscsegment":"53000000","description":"Apparel & Luggage & Personal Care Products","updatedTotalCost":11157130.00,"totalCount":55,"totalCost":11157130.00,"typeofapp":3},{"unspscsegment":"46000000","description":"Defense & Law Enforcement & Security & Safety Equipment & Supplies","updatedTotalCost":7407310.00,"totalCount":55,"totalCost":7407310.00,"typeofapp":3},{"unspscsegment":"27000000","description":"Tools & General Machinery","updatedTotalCost":3281900.00,"totalCount":55,"totalCost":3281900.00,"typeofapp":3},{"unspscsegment":"24000000","description":"Material Handling & Conditioning & Storage Machinery & their Accessories & Supplies","updatedTotalCost":2201100.00,"totalCount":55,"totalCost":2201100.00,"typeofapp":3},{"unspscsegment":"40000000","description":"Distribution & Conditioning Systems & Equipment & Components","updatedTotalCost":3654250.00,"totalCount":55,"totalCost":3654250.00,"typeofapp":3},{"unspscsegment":"21000000","description":"Farming & Fishing & Forestry & Wildlife Machinery & Accessories","updatedTotalCost":2040900.00,"totalCount":55,"totalCost":2040900.00,"typeofapp":3}]}}

