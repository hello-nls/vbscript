' source: https://community.nortridge.com/t/disable-transactions/1857/

    lnGrp = nlsApp.GetField("LOAN_GROUP")
    if (lnGrp = "SIMPLE INTEREST") then
          msgbox ("sorry, can't manually post any transaction on this loan.")
          nlsApp.break()
    end if
