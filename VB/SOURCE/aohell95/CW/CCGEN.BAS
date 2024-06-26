  Global Const KEY_RETURN = &HD
  Global banc(7000)
  Global valid As Integer
  Global TheCard As Integer
  Dim shh(17)
  Global decrd(17)
  Global cantdoextrap
  Global cardstring(20)


  Function checkvalid (TheCreditCard As String) As Integer

    On Error GoTo assmunch
  'reinitialize
    sum = 0
    For jjj = 1 To 16
      decrd(jjj) = 0
    Next jjj
  'starting vars
    qqq$ = TheCreditCard
    x = 1
  'begin
    For kount = 1 To 16
      If Mid$(qqq$, x, 1) = " " Then
        x = x + 1
        kount = kount - 1
      Else
        cardstring(kount) = Mid(qqq$, x, 1)
        x = x + 1
      End If
    Next kount
    For xr = 1 To 16
      fukaol = Val(cardstring(xr))
      decrd(xr) = fukaol
    Next xr
    For cc = 1 To 15 Step 2
      decrd(cc) = decrd(cc) * 2
      If decrd(cc) > 9 Then decrd(cc) = decrd(cc) - 9
    Next cc
    For bbb = 1 To 16
      sum = sum + decrd(bbb)
    Next bbb
    xre = sum Mod 10
    If xre = 0 Then
      checkvalid = True
    Else
      checkvalid = False
    End If
    Exit Function
    assmunch:
    If Err = 13 Then
      MsgBox "Please Enter The Card Number With No Letters or Hyphens."
    Else
      MsgBox "Something Weird Happened"
    End If
    Exit Function
  End Function

  Sub loadbanks ()
    banc(4013) = "4013 - Citibank"
    banc(4019) = "4019 - Bank of America"
    banc(4024) = "4024 - Bank of America"
    banc(4027) = "4027 - Rockwell Federal Credit Union"
    banc(4032) = "4032 - Household Bank"
    banc(4052) = "4052 - First Cincinatti"
    banc(4060) = "4060 - Associates National Bank"
    banc(4070) = "4070 - Security Pacific"
    banc(4071) = "4071 - Colonial National Bank"
    banc(4094) = "4094 - A.M.C. Federal Credit Union"
    banc(4113) = "4113 - Valley National Bank"
    banc(4114) = "4114 - Chemical Bank"
    banc(4121) = "4121 - Pennsylvania State Employees Credit Union"
    banc(4122) = "4122 - Union Trust"
    banc(4128) = "4128 - Citibank"
    banc(4131) = "4131 - State Street Bank"
    banc(4215) = "4125 - Marine Midland"
    banc(4225) = "4225 - Chase Manhatten Bank"
    banc(4226) = "4226 - Chase Manhatten Bank"
    banc(4231) = "4231 - Chase Lincoln First Classic"
    banc(4232) = "4232 - Chase Lincoln First Classic"
    banc(4239) = "4239 - Corestates"
    banc(4241) = "4241 - National Westminester Bank"
    banc(4250) = "4250 - First Chicago Bank"
    banc(4253) = "4253 - Consumers Edge"
    banc(4254) = "4254 - Security First"
    banc(4271) = "4271 - Citibank Preferred"
    banc(4301) = "4301 - Monogram Bank"
    banc(4302) = "4302 - H.H.B.C."
    banc(4310) = "4310 - B.C.F.U."
    banc(4311) = "4311 - First National Bank of Louisville"
    banc(4317) = "4317 - Gold Dome"
    banc(4327) = "4327 - First Atlanta"
    banc(4332) = "4332 - First American Bank"
    banc(4339) = "4339 - Primerica Bank"
    banc(4342) = "4342 - N.C.M.B. / Nations Bank"
    banc(4356) = "4356 - National Bank of Delaware"
    banc(4368) = "4368 - National West"
    banc(4387) = "4387 - Bank One"
    banc(4388) = "4388 - First Signature Bank & Trust"
    banc(4401) = "4401 - Gary-Wheaton Bank"
    banc(4413) = "4413 - Firstier Bank Lincoln"
    banc(4418) = "4418 - Bank of Omaha"
    banc(4421) = "4421 - Indiana National Bank"
    banc(4428) = "4428 - Bank of Hoven"
    banc(4436) = "4436 - Security Bank and Trust"
    banc(4443) = "4443 - Merril Lynch Bank/Trust"
    banc(4447) = "4447 - AmeriTrust"
    banc(4452) = "4452 - Empire Affiliates Federal Credit Union"
    banc(4498) = "4498 - Republic Savings"
    banc(4502) = "4502 - C.I.B.C."
    banc(4503) = "4503 - Canadian Imperial Bank"
    banc(4506) = "4506 - Belgium A.S.L.K"
    banc(4510) = "4510 - Royal Bank of Canada"
    banc(4520) = "4520 - Toronto Dominion of Canada"
    banc(4537) = "4537 - Bank of Nova Scotia"
    banc(4538) = "4538 - Bank of Nova Scotia"
    banc(4539) = "4539 - Barclays"
    banc(4544) = "4544 - T.S.B. Bank"
    banc(4556) = "4556 - Citibank"
    banc(4564) = "4564 - Bank of Queensland"
    banc(4673) = "4673 - First Card"
    banc(4707) = "4707 - Tompkins County Trust"
    banc(4719) = "4719 - Rocky Mountain"
    banc(4721) = "4721 - First Security"
    banc(4722) = "4722 - West Bank"
    banc(4726) = "4726 - Wells Fargo"
    banc(4783) = "4783 - AT&T's Universal Card"
    banc(4784) = "4784 - AT&T's Universal Card"
    banc(4800) = "4800 - M.B.N.A. North America"
    banc(4811) = "4811 - Bank of Hawaii"
    banc(4819) = "4819 - Macom Federal Crdit Union"
    banc(4820) = "4820 - IBM Mid America Federal Credit Union"
    banc(4833) = "4833 - U.S. Bank"
    banc(4842) = "4842 - Security Pacific Washington"
    banc(4897) = "4897 - Village Bank of Chicago"
    banc(4921) = "4921 - Hong Kong/National Bank"
    banc(5100) = "5100 - Southwestern States Bankcard Association"
    banc(5110) = "5110 - Universal Travel Voucher"
    banc(5120) = "5120 - Western States Bankcard Association"
    banc(5130) = "5130 - Eurocard France"
    banc(5140) = "5140 - Mountain States Bankcard Association"
    banc(5150) = "5150 - Credit Systems, Inc."
    banc(5160) = "5160 - Westpac Banking Corporation"
    banc(5170) = "5170 - Midamerica Bankcard Association"
    banc(5172) = "5172 - First Bank Card Center"
    banc(5180) = "5180 - Computer Communications of America"
    banc(5190) = "5190 - Bank of Montreal"
    banc(5191) = "5191 - Bank of Montreal"
    banc(5201) = "5201 - Mellon Bank, N.A."
    banc(5202) = "5202 - Central Trust Company, N.A."
    banc(5204) = "5204 - Security Pacific National Bank"
    banc(5205) = "5205 - Promocion Y Operacion, S.A."
    banc(5206) = "5206 - Banco Nacional de Mexico"
    banc(5207) = "5207 - New England Bankcard Association, Inc."
    banc(5208) = "5208 - Million Card Service Co., Ltd."
    banc(5209) = "5209 - The Citizens & Southern National Bank"
    banc(5210) = "5210 - Kokunai Shinpan Company, Ltd."
    banc(5211) = "5211 - Chemical Bank Delaware"
    banc(5212) = "5212 - F.C.C. National Bank"
    banc(5213) = "5213 - The Bankcard Association, Inc."
    banc(5215) = "5215 - Marine Midland Bank, N.A."
    banc(5216) = "5216 - Old Kent Bank & Trust Co."
    banc(5217) = "5217 - Union Trust"
    banc(5218) = "5218 - Citibank, N.A."
    banc(5219) = "5219 - Central Finance Co., Ltd."
    banc(5220) = "5220 - Sovran Bank/Central South"
    banc(5221) = "5221 - Standard Bank of South Africa, Ltd."
    banc(5222) = "5222 - Security Bank & Trust Company"
    banc(5223) = "5223 - Trustmark National Bank"
    banc(5224) = "5224 - Midland Bank"
    banc(5225) = "5225 - First Pennsylvania Bank, N.A."
    banc(5226) = "5226 - Eurocard Ab"
    banc(5227) = "5227 - Rocky Mountain Bankcard System, Inc"
    banc(5228) = "5228 - First Union National Bank of North Carolina"
    banc(5229) = "5229 - Sunwest Bank of Albuquerque, N.A."
    banc(5230) = "5230 - Harris Trust & Savings Bank"
    banc(5231) = "5231 - Badische Beamtenbank EG"
    banc(5232) = "5232 - Eurocard Deutschland"
    banc(5233) = "5233 - Bancsystems Association, Inc."
    banc(5234) = "5234 - Citibank <Arizona>"
    banc(5235) = "5235 - Finanacial Transaction Systems, Inc."
    banc(5236) = "5236 - First Tennessee Bank, N.A."

    banc(5333) = "5333 - Ohio National Bank"
    banc(5351) = "5351 - Provident National Bank"
    banc(6013) = "6013 - MBNA"

  End Sub

