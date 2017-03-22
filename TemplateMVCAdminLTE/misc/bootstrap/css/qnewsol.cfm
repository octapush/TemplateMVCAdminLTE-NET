<!------------------------------------------------------------------------------
APPLICATION......: SunFishERP 51
FID..............: -
FUID/SELACTIONID.: -
FILENAME.........: qnewsol.cfm
================================================================================
CREATED BY.......: - ??
CREATED DATE.....: - ??
================================================================================
DESCRIPTION......:  ??
================================================================================
REVISION.........: /* 23 September 2010 - randytia */
.................: Menambahkan check RFQ Number
------------------------------------------------------------------------------->
<!--- <cfset Payment_Term = "#FORM['cboTermMenu']#">
<cfif "#Payment_Term#" EQ "P001">
  <cfset miles = 0>
  <cfloop index="i" from="1" to="#FORM['hitMenu']#">
      <cfset miles = isDefined(val(FORM['miles_#i#'])) ? 0 : val(FORM['miles_#i#']) >
      <cfquery name="qNewPayment" datasource="#iif(isDefined('DSN'),'DSN','Attributes.DSN')#">
        INSERT INTO dbo.TAccPayment_Term(
            Payment_Term
            ,Payment_Term_Type
            ,Percentage
            ,Quotation_Number
            ,SO_Number
            ,Invoice_Number
            ,Invoice_Paid
            ,add_by
            ,add_dt
          ) 
        VALUES 
          (
              '#FORM['cboTermMenu']#',
              'NET_30',
              '#miles#', --FOR PRESENTAGE
              'QUOTATION_xx#i#',
              'SO_xx#i#',
              'INV_XX#i#',
              '1',
              '#COOKIE.CKSATRIADEVID#',
              GETDATE()
          )
      </cfquery>
      <cfoutput>"#FORM['miles_#i#']#" SUBMILES => "#strDocPattern#"</cfoutput>
  </cfloop>
    <cfoutput>Ke Generate #miles#</cfoutput>
  <cfelse>
    <cfoutput>Tidak Ke Generate #miles#</cfoutput>
</cfif> --->
<!--- <cfabort> --->

<cfoutput>
<!--- UPDATE dbo.TAccPayment_Term
                  SET
                      dbo.TAccPayment_Term.is_active = '0'
                  WHERE dbo.TAccPayment_Term.Term_Id IN(#FORM['PaymentKey']#)
<cfabort> --->
<cfparam name="url.taskRevisi" type="string" default="" />
<cfif url.taskRevisi EQ "new">
  <cfset taskRevisi = url.taskRevisi>
<cfelse>
  <cfset taskRevisi = "FALSE">
</cfif>
TASK REVISI #url.taskRevisi#
<!--- <cfabort> --->
  <cfif listfindnocase(menu,"purchase")>
    <cfset submenu = "purchase">
    <cfelse>
    <cfset submenu = "sales">
  </cfif>
  <cfif listfindnocase(menu,"sales")>
    <cfset varSecAccess = REQUEST.SFSecAccess.SecAccessFile(FILEACCESSCODE="ERSTD0803308", BACKURL="#Application.stApp.Web_Path[VST_IDX]#/#Application.stApp.Home_URL[VST_IDX]#/index.cfm?selListItem=1&menu=0")>
    <cfelse>
    <cfset varSecAccess = REQUEST.SFSecAccess.SecAccessFile(FILEACCESSCODE="ERSTD0798604", BACKURL="#Application.stApp.Web_Path[VST_IDX]#/#Application.stApp.Home_URL[VST_IDX]#/index.cfm?selListItem=1&menu=0")>
  </cfif>
  <cfset objDummy = createObject("component", "#Application.ComponentPath#.sunfisherp.utility.cfdummy")>
  <cfset LANGUAGELIST="StepNotAvailable,set,careerHistory,IsNotDefined,SuccessInserted,eHRMUpdatedSuccess,TaxConverter,CurrencyConverter,QuotationNo">
  <CF_DO_V25_MULTILANGUAGE MESSAGEIDLIST="#LanguageList#">
  <cftry>
    <cfquery datasource="#request.dsn#">
  alter table TACCQuotation_ItemColor add config_order int
  </cfquery>
    <cfcatch>
    </cfcatch>
  </cftry>
  <cfparam name="form.hdnDocRelation" default=""/>
  <cfif Trim(selRFQ) neq "" and Task neq "edit" AND NOT isDefined("SourceDoc")>
    <cffunction name="fntCheckRFQ" access="private" returntype="boolean">
      <cfargument name="varRFQNumber" required="yes" type="string">
      <cfparam name="local.blnRFQExist" default="true" type="boolean">
      <cfparam name="local.strErrMsg" default="" type="string">
      <cftry>
        <cfquery name="qCheckRFQ" datasource="#REQUEST.DSN#">
            Select RFQ_Code
            from TAccRFQ_header, TAccount
            where TAccRFQ_header.Account_ID = TAccount.Account_ID
            <cfif submenu eq "sales">
              AND RFQ_type = 'sales'
            <cfelse>
              AND RFQ_type = 'purchase'
              AND Approval_Status = 3
            </cfif>
              AND RFQ_category = '#selCatType#'
              AND TAccRFQ_header.Company_ID = #COOKIE.COMPANYID#
            AND isnull(TAccRFQ_header.EXPIRED,0) = 0
            And RFQ_Code =  '#Trim(selRFQ)#'
        </cfquery>
        <cfif qCheckRFQ.recordcount NEQ 1>
          <cfset local.blnRFQExist = false>
        </cfif>
        <cfreturn local.blnRFQExist>
        <cfcatch type="any">
          <cfset local.strErrMsg = "Message : " & cfcatch.Message & "<br />Detail : " & cfcatch.Detail>
          #objDummy.cferror(ErrorText: local.strErrMsg)#
        </cfcatch>
      </cftry>
    </cffunction>
    <cfset local.blnChcek = fntCheckRFQ(varRFQNumber: FORM['selRFQ'])>
    <cfif local.blnChcek eq false>
      <cfset local.strErrMsg = "Message : Document number [" & FORM['selRFQ'] & "] is not found!<br />Detail : Please choose existing document">
      #objDummy.cferror(ErrorText: local.strErrMsg)#
      <cfabort>
    </cfif>
  </cfif>
  <cfif isDefined("SourceDoc")>
    <!--- Original Source --->
    <cfquery name="qSource" datasource="#REQUEST.DSN#">
      SELECT isNULL(Original_Doc,Quotation_Number) AS Original_Doc, isLead,Lead_Number FROM TAccQuotation_Header WHERE Quotation_Number = '#SourceDoc#'
    </cfquery>
    <cfset Original_Doc = qSource.Original_Doc>
    <cfset ORIisLead = qSource.isLead>
    <cfset ORILead_Number = qSource.Lead_Number>
    <!--- Revision Number --->
    <cfquery name="qRev" datasource="#REQUEST.DSN#">
      SELECT isNULL(MAX(Revision),0) + 1 AS Revision FROM TAccQuotation_Header WHERE Original_Doc = '#Original_Doc#'
    </cfquery>
    <cfif qRev.RecordCount GT 0>
      <cfset Revision = qRev.Revision>
      <cfelse>
      <cfset Revision = 1>
    </cfif>
  </cfif>
  <cftransaction>
    <!--- Insert New Item CRF51014-14497 YS 20141008 --->
    <CF_DO_CS_UPLOADITEM TYPE="QUERY" MENU="#MENU#">
    <!--- End CRF51014-14497 --->
    <cfset DocumentDate = #txtTgl#>
    <cfinclude template="#Application.stApp.CFWeb_Path[1]#/include/lockperiod/locktransaction.cfm">
    <cfif task neq "Edit">
      <cfif menu IS "sales">
        <cfset strDocPattern = "CustQuotation">
        <cfelse>
        <cfset strDocPattern = "VendorQuotation">
      </cfif>
      <cfif menu IS "sales">
        <cfif isdefined('Revision') and Revision GT 0>
          <CF_DO_CS_ACCDOCUMENTNO TableName="TAccPattern" DocumentType="#strDocPattern#" DocumentNo="Quotation_Number" Type="value" CompanyID="#Cookie.CompanyID#" LocationID="#COOKIE.LOCATION_ID#" TrxNo="Trans" CustomNo="#Original_Doc#" Postfix="REV" PostfixCounter="#Revision#">
          <cfelse>
          <CF_DO_CS_ACCDOCUMENTNO TableName="TAccPattern" DocumentType="#strDocPattern#" DocumentNo="Quotation_Number" Type="value" CompanyID="#Cookie.CompanyID#" LocationID="#COOKIE.LOCATION_ID#" Prefix="#ListLast(selDevision,'|')#-#selSalesCode#" PrefixOrder="2" TrxNo="Trans" DivisionID="#ListLast(selDevision,'|')#">
        </cfif>
        <cfelse>
        <CF_DO_V30_ACCDOCUMENTNO TableName="TAccPattern" DocumentType="#strDocPattern#" DocumentNo="Quotation_Number" Type="value" CompanyID="#Cookie.CompanyID#" LocationID="#COOKIE.LOCATION_ID#" TrxNo="Trans">
      </cfif>
    </cfif>
    <!--- <cfset Rate = #val(replace(evaluate("txtCurr_#selCurrency#"),",","","ALL"))#> --->
    <cfif selCurrency EQ Cookie.currencyid>
      <cfset Rate = 1>
      <cfelse>
      <cfset Rate = replace(evaluate("txtCurr_#selCurrency#"),",","","ALL")>
    </cfif>
    <cfset InvoiceAmount  = replace(txtTotAmount,",","","ALL")>
    <!--- START UPDATE BY BEDU --->
  <cfset BaseInvoiceAmount  = InvoiceAmount * Rate>
    <cfset BaseInvoiceAmount  = INT(BaseInvoiceAmount)>
    <!--- END UPDATE BY BEDU --->
    <cfif cookie.currencyid2 neq 0>
      <!--- <cfset Rate2 = #val(replace(evaluate("txtCurr2_#selCurrency#"),",","","ALL"))#> --->
      <cfset Rate2 = replace(evaluate("txtCurr2_#selCurrency#"),",","","ALL")>
      <cfset BaseInvoiceAmount2  = InvoiceAmount * Rate2>
    </cfif>
    <!--- <cfset txtTotTaxConv  = #val(replace(txtTotTaxConv,",","","ALL"))#>  --->
    <cfset txtTotTaxConv  = replace(txtTotTaxConv,",","","ALL")>
    <!--- <cfset txtTotTaxConv_Base  = txtTotTaxConv * #val(replace( evaluate("txtTax_#seltaxCurrency#") ,",","","ALL"))#> --->
    <cfif seltaxCurrency EQ Cookie.currencyid>
      <cfset txtTotTaxConv_Base = txtTotTaxConv>
      <cfelse>
      <cfset txtTotTaxConv_Base  = txtTotTaxConv * replace( evaluate("txtTax_#seltaxCurrency#") ,",","","ALL")>
    </cfif>
    <cfif cookie.currencyid2 neq 0>
      <!--- <cfset txtTotTaxConv_Base2  = txtTotTaxConv * #val(replace( evaluate("txtTax2_#seltaxCurrency#") ,",","","ALL"))#> --->
      <cfset txtTotTaxConv_Base2  = txtTotTaxConv * replace( evaluate("txtTax2_#seltaxCurrency#") ,",","","ALL")>
    </cfif>
    <cfset txtTotDeductConv  = #val(replace(txtTotDeductConv,",","","ALL"))#>
    <cfif seltaxCurrency EQ Cookie.currencyid>
      <cfset txtTotDeductConv_Base = txtTotDeductConv>
      <cfelse>
      <cfset txtTotDeductConv_Base  = txtTotDeductConv * #val(replace( evaluate("txtTax_#seltaxCurrency#") ,",","","ALL"))#>
      <!---tidak dipakai ya?--->
    </cfif>
    <cfif cookie.currencyid2 neq 0>
      <cfset txtTotDeductConv_Base2  = txtTotDeductConv * #val(replace( evaluate("txtTax2_#seltaxCurrency#") ,",","","ALL"))#>
    </cfif>
    <cfset CurrencyRateList="#cookie.currencyid#|1">
    <cfset TaxRateList="#cookie.currencyid#|1">
    <cfset AmountCurrency="#cookie.currencyid#">
    <cfset TaxCurrency="#cookie.currencyid#">
    <cfif cookie.currencyid2 neq 0>
      <cfset CurrencyRateList2="#cookie.currencyid2#|1">
      <cfset TaxRateList2="#cookie.currencyid2#|1">
      <cfset AmountCurrency2="#cookie.currencyid2#">
      <cfset TaxCurrency2="#cookie.currencyid2#">
    </cfif>
    <cfloop list="#lstCurrency#" index="ListAwal" delimiters=";">
      <cfset TypeofTransaction=listgetat(ListAwal,1,"|")>
      <cfset Currency=listgetat(ListAwal,2,"|")>
      <cfif TypeofTransaction eq "Amount">
        <cfif listfindnocase(AmountCurrency,Currency) eq "0">
          <cfset converter = #val(replace(evaluate("txtCurr_#Currency#"),",","","ALL"))#>
          <cfset CurrencyRateList=listappend(CurrencyRateList,"#Currency#|#precisionevaluate(converter)#",";")>
          <cfset AmountCurrency=listappend(AmountCurrency,Currency)>
          <cfif converter eq 0>
            <script>
              alert('#DO_VAR["CurrencyConverter"]# #Currency# #DO_VAR["IsNotDefined"]#');
              history.back();
            </script>
            <cfabort>
          </cfif>
        </cfif>
      </cfif>
      <cfif TypeofTransaction eq "Tax">
        <cfif listfindnocase(TaxCurrency,Currency) eq "0">
          <cfset converter = #val(replace(evaluate("txtTax_#Currency#"),",","","ALL"))#>
          <cfset TaxRateList=listappend(TaxRateList,"#Currency#|#precisionevaluate(converter)#",";")>
          <cfset TaxCurrency=listappend(TaxCurrency,Currency)>
          <cfif converter eq 0>
            <script>
              alert('#DO_VAR["TaxConverter"]# #Currency# #DO_VAR["IsNotDefined"]#');
              history.back();
            </script>
            <cfabort>
          </cfif>
        </cfif>
      </cfif>
      <!---untk dualbase--->
      <cfif cookie.currencyid2 neq 0>
        <cfif TypeofTransaction eq "Amount">
          <cfif listfindnocase(AmountCurrency2,Currency) eq "0">
            <!--- <cfset converter = #val(replace(evaluate("txtCurr2_#Currency#"),",","","ALL"))#> --->
            <cfset converter = #replace(evaluate("txtCurr2_#Currency#"),",","","ALL")#>
            <cfset CurrencyRateList2=listappend(CurrencyRateList2,"#Currency#|#precisionevaluate(converter)#",";")>
            <cfset AmountCurrency2=listappend(AmountCurrency2,Currency)>
            <cfif converter eq 0>
              <script>
                  alert('#DO_VAR["CurrencyConverter"]# #Currency# #DO_VAR["IsNotDefined"]#');
                  //history.back();
                </script>
              <cfabort>
            </cfif>
          </cfif>
        </cfif>
        <cfif TypeofTransaction eq "Tax">
          <cfif listfindnocase(TaxCurrency2,Currency) eq "0">
            <cfset converter = #val(replace(evaluate("txtTax2_#Currency#"),",","","ALL"))#>
            <cfset TaxRateList2=listappend(TaxRateList2,"#Currency#|#precisionevaluate(converter)#",";")>
            <cfset TaxCurrency2=listappend(TaxCurrency2,Currency)>
            <cfif converter eq 0>
              <script>
                  alert('#DO_VAR["TaxConverter"]# #Currency# #DO_VAR["IsNotDefined"]#');
                  history.back();
                </script>
              <cfabort>
            </cfif>
          </cfif>
        </cfif>
      </cfif>
      <!---END untk dualbase--->
    </cfloop>
    <!---<cfif isDefined ("form.rbTypedoc") and form.rbTypedoc eq 0 and form.selRFQ neq 0>
    <cfquery name="qUpdate" datasource="#iif(isDefined('DSN'),'DSN','Attributes.DSN')#">
      update taccrfq_header set rfq_status = 3
      where rfq_code = '#form.selRFQ#'
    </cfquery>
  </cfif>--->
    <!---MMY 20150112, tidak perlu update rfq_status--->
    <cfquery name="qaddress" datasource="#iif(isDefined('DSN'),'DSN','Attributes.DSN')#">
    select account_address1 from taccount where account_id='#txtCustCode#'
  </cfquery>
    <!---Angries--->
    <cfif isdefined("cboTerms") and #cboTerms# neq "">
      <cfquery name="qGetTypePaymentTerm" datasource="#REQUEST.DSN#">
          select Top_type from TAccTermOfPayment_Header
          where TOP_Code = '#cboTerms#'
      </cfquery>
    </cfif>

    <!--- EDIT  BY ADAM
    /+++++++++++++++++++++++++++++++++++++++++++++++/
    / AUTHOR        : ADAM SUMARNA                  /
    / DATE          : 5-09-2016                     /
    / DESCRIPTION   : QUERY QUOTATION PAYMENT TERM  /
    /+++++++++++++++++++++++++++++++++++++++++++++++/
    --->

    <cfif task EQ "Save" AND FORM['newQuotationNumber'] EQ "FALSE" >
        <cfset MSG = "">
        <cfset updatePayment = "#FORM['stTask']#">
        <cfset Payment_Term = "#FORM['cboTermMenu']#">
        <cfset tempPayment = "#FORM['tempPayment']#">

        <cfif taskRevisi NEQ "revisi">
            <cfif updatePayment EQ "updatePayment">
              <cfquery name="qRemoveQuo" datasource="#REQUEST.DSN#">
                DECLARE @QuotationNumber AS VARCHAR(50);
                DECLARE @PaymentTerm AS VARCHAR(10);
                SET @QuotationNumber = '#quotation_number#';
                SET @PaymentTerm = '#tempPayment#';
                IF EXISTS
                (
                    SELECT TOP 1 tpt.Quotation_Number
                    FROM dbo.TAccPayment_Term tpt
                    WHERE LEFT(tpt.Quotation_Number,13) = LEFT('#quotation_number#',13) OR tpt.Quotation_Number = '#quotation_number#'
                          AND tpt.Payment_Term = @PaymentTerm
                          AND tpt.SO_Number IS NULL 
                          AND tpt.is_revise = 0
                )
                    BEGIN
                        UPDATE dbo.TAccPayment_Term SET dbo.TAccPayment_Term.is_revise=1
                        WHERE dbo.TAccPayment_Term.Payment_Term = @PaymentTerm 
                            AND LEFT(dbo.TAccPayment_Term.Quotation_Number,13) = LEFT('#quotation_number#',13) OR dbo.TAccPayment_Term.Quotation_Number = '#quotation_number#'
                              AND dbo.TAccPayment_Term.SO_Number IS NULL;
                        SELECT 'TRUE' AS MSG  
                    END;
                ELSE
                    BEGIN
                        SELECT 'FALSE' AS MSG
                    END;
              </cfquery>
              <cfset MSG = qRemoveQuo.MSG>
            </cfif>
            <cfset MSG = "FALSE">
        </cfif>
          <!--- 
          DATA AKAN DI HAPUS BILA DATA PAYMENT TERM INI SUDAH ADA DAN SO NUMBER MASI BELUM DI BUAT, DAN PAYMENT TERM AKAN DI INPUT BARU 
          --->

          <cfif MSG EQ "TRUE" AND Payment_Term NEQ tempPayment AND updatePayment EQ "updatePayment">
            <cfif Payment_Term EQ "P001">
                <cfif FORM['PaymentKey'] NEQ "">
                  <cfquery name="qUpdateNet30WithKey" datasource="#REQUEST.DSN#">
                    UPDATE dbo.TAccPayment_Term
                    SET
                        dbo.TAccPayment_Term.is_active = '0'
                    WHERE dbo.TAccPayment_Term.Term_Id IN(#FORM['PaymentKey']#)
                  </cfquery>
                </cfif>

                <cfquery name="qUpdateNet30" datasource="#REQUEST.DSN#">
                  INSERT INTO dbo.TAccPayment_Term
                  (Payment_Term,Payment_Term_Type,Percentage,Quotation_Number,SO_Number,Invoice_Number,Invoice_Paid,add_by,add_dt,mod_by,mod_dt,referencen_old,status_inv,is_active)
                  VALUES
                  (
                      '#Payment_Term#', 
                      'NET30', 
                      100.00, 
                      '#quotation_number#', 
                      NULL, 
                      NULL,
                      0, 
                      '#COOKIE.CKSATRIADEVID#',
                      GETDATE(),
                      NULL, 
                      NULL, 
                      NULL, 
                      0, 
                      1 
                  )
                </cfquery>
            </cfif>
            <cfif Payment_Term EQ "P002">
                  <cfset eDp        = "#FORM['eDp']#">
                  <cfset eBp        = "#FORM['eBp']#">
                  <cfset selectDp   = "#FORM['selectDp']#">
                  <cfset selectBp   = "#FORM['selectBp']#">
                  <cfset idDp       = "#FORM['idEdp']#">
                  <cfset idBp       = "#FORM['idEbp']#">

                  <cfif FORM['PaymentKey'] NEQ "">
                    <cfquery name="qUpdateNet30WithKey" datasource="#REQUEST.DSN#">
                      UPDATE dbo.TAccPayment_Term
                      SET
                          dbo.TAccPayment_Term.is_active = '0'
                      WHERE dbo.TAccPayment_Term.Term_Id IN(#FORM['PaymentKey']#)
                    </cfquery>
                  </cfif>

                  <cfquery name="qInsertDp" datasource="#REQUEST.DSN#">
                    INSERT INTO dbo.TAccPayment_Term
                    (Payment_Term,Payment_Term_Type,Percentage,Quotation_Number,SO_Number,Invoice_Number,Invoice_Paid,add_by,add_dt,mod_by,mod_dt,referencen_old,status_inv,is_active)
                    VALUES
                    (
                        '#Payment_Term#', 
                        'DP002', 
                        '#eDp#', 
                        '#quotation_number#', 
                        NULL, 
                        NULL,
                        0, 
                        '#COOKIE.CKSATRIADEVID#',
                        GETDATE(),
                        NULL, 
                        NULL, 
                        NULL, 
                        0, 
                        1 
                    ),
                    (
                        '#Payment_Term#', 
                        'BP003', 
                        '#eBp#', 
                        '#quotation_number#', 
                        NULL, 
                        NULL,
                        0, 
                        '#COOKIE.CKSATRIADEVID#',
                        GETDATE(),
                        NULL, 
                        NULL, 
                        NULL, 
                        0, 
                        1 
                    )
                  </cfquery>
            </cfif>
            <cfif Payment_Term EQ "P003">
                <cfif FORM['PaymentKey'] NEQ "">
                  <cfquery name="qUpdateNet30WithKey" datasource="#REQUEST.DSN#">
                    UPDATE dbo.TAccPayment_Term
                    SET
                        dbo.TAccPayment_Term.is_active = '0'
                    WHERE dbo.TAccPayment_Term.Term_Id IN(#FORM['PaymentKey']#)
                  </cfquery>
                </cfif>

                <cfloop index="i" from="1" to="#FORM['hitMenu']#">
                  <cfset miles = FORM['miles_#i#']>
                  <cfset subMiles = FORM['subMiles_#i#']>
                  <cfquery name="qInsertMiles" datasource="#REQUEST.DSN#">
                      INSERT INTO dbo.TAccPayment_Term
                      (Payment_Term,Payment_Term_Type,Percentage,Quotation_Number,SO_Number,Invoice_Number,Invoice_Paid,add_by,add_dt,mod_by,mod_dt,referencen_old,status_inv,is_active)
                      VALUES
                      (
                          '#Payment_Term#', 
                          '#subMiles#', 
                          '#miles#', 
                          '#quotation_number#', 
                          NULL, 
                          NULL,
                          0, 
                          '#COOKIE.CKSATRIADEVID#',
                          GETDATE(),
                          NULL, 
                          NULL, 
                          NULL, 
                          0, 
                          1 
                      )
                  </cfquery>
                </cfloop>
              </cfif>
          
          <!--- END qRemoveQuo--->
          
          <cfelse>
            <cfif Payment_Term EQ "P001">

              <cfif FORM['PaymentKey'] NEQ "">
                <cfquery name="qUpdateNet30WithKey" datasource="#REQUEST.DSN#">
                  UPDATE dbo.TAccPayment_Term
                  SET
                      dbo.TAccPayment_Term.is_active = '0'
                  WHERE dbo.TAccPayment_Term.Term_Id IN(#FORM['PaymentKey']#)
                </cfquery>
              </cfif>

              <cfquery name="qUpdateNet30" datasource="#REQUEST.DSN#">
                 INSERT INTO dbo.TAccPayment_Term
                  (Payment_Term,Payment_Term_Type,Percentage,Quotation_Number,SO_Number,Invoice_Number,Invoice_Paid,add_by,add_dt,mod_by,mod_dt,referencen_old,status_inv,is_active)
                  VALUES
                  (
                      '#Payment_Term#', 
                      'NET30', 
                      100.00, 
                      '#quotation_number#', 
                      NULL, 
                      NULL,
                      0, 
                      '#COOKIE.CKSATRIADEVID#',
                      GETDATE(),
                      NULL, 
                      NULL, 
                      NULL, 
                      0, 
                      1 
                  )
              </cfquery>
            </cfif>
            <cfif Payment_Term EQ "P002">
                <cfset eDp        = "#FORM['eDp']#">
                <cfset eBp        = "#FORM['eBp']#">
                <cfset selectDp   = "#FORM['selectDp']#">
                <cfset selectBp   = "#FORM['selectBp']#">
                <cfset idDp       = "#FORM['idEdp']#">
                <cfset idBp       = "#FORM['idEbp']#">

                <cfif FORM['PaymentKey'] NEQ "">
                  <cfquery name="qUpdateDpWithKey" datasource="#REQUEST.DSN#">
                    UPDATE dbo.TAccPayment_Term
                    SET
                        dbo.TAccPayment_Term.is_active = '0'
                    WHERE dbo.TAccPayment_Term.Term_Id IN(#FORM['PaymentKey']#)
                  </cfquery>
                </cfif>

                <cfquery name="qUpdateDp" datasource="#REQUEST.DSN#">
                  INSERT INTO dbo.TAccPayment_Term
                    (Payment_Term,Payment_Term_Type,Percentage,Quotation_Number,SO_Number,Invoice_Number,Invoice_Paid,add_by,add_dt,mod_by,mod_dt,referencen_old,status_inv,is_active)
                    VALUES
                    (
                        '#Payment_Term#', 
                        'DP002', 
                        '#eDp#', 
                        '#quotation_number#', 
                        NULL, 
                        NULL,
                        0, 
                        '#COOKIE.CKSATRIADEVID#',
                        GETDATE(),
                        NULL, 
                        NULL, 
                        NULL, 
                        0, 
                        1 
                    ),
                    (
                        '#Payment_Term#', 
                        'BP003', 
                        '#eBp#', 
                        '#quotation_number#', 
                        NULL, 
                        NULL,
                        0, 
                        '#COOKIE.CKSATRIADEVID#',
                        GETDATE(),
                        NULL, 
                        NULL, 
                        NULL, 
                        0, 
                        1 
                    )
                </cfquery>
            </cfif>
            <cfif Payment_Term EQ "P003">

              <cfif FORM['PaymentKey'] NEQ "">
                <cfquery name="gUpdateMilesWithKey" datasource="#REQUEST.DSN#">
                  UPDATE dbo.TAccPayment_Term
                  SET
                      dbo.TAccPayment_Term.is_active = '0'
                  WHERE dbo.TAccPayment_Term.Term_Id IN(#FORM['PaymentKey']#)
                </cfquery>
              </cfif>

              <cfloop index="i" from="1" to="#FORM['hitMenu']#">
                <cfset miles = FORM['miles_#i#']>
                <cfset subMiles = FORM['subMiles_#i#']>
                <cfquery name="qInsertMiles" datasource="#REQUEST.DSN#">
                    INSERT INTO dbo.TAccPayment_Term
                    (Payment_Term,Payment_Term_Type,Percentage,Quotation_Number,SO_Number,Invoice_Number,Invoice_Paid,add_by,add_dt,mod_by,mod_dt,referencen_old,status_inv,is_active)
                    VALUES
                    (
                        '#Payment_Term#', 
                        '#subMiles#', 
                        '#miles#', 
                        '#quotation_number#', 
                        NULL, 
                        NULL,
                        0, 
                        '#COOKIE.CKSATRIADEVID#',
                        GETDATE(),
                        NULL, 
                        NULL, 
                        NULL, 
                        0, 
                        1 
                    )
                </cfquery>
              </cfloop>
            </cfif>
          <!--- END SCALAR QUERY INFO --->
          </cfif>
        <!--- </cfif> --->
    </cfif>
    <!--- END --->
    <!---end--->

    <!---MMY 20141028, Document Checklist--->
    <CF_DO_CS_DOCCHECKLIST_TRX MODULE_CODE="SQT" TRX_NUMBER="#Quotation_Number#">
    <cfif Task eq "save">
      <!---<CF_DO_V30_ACCDOCUMENTNO TableName="TAccPattern" DocumentType="CustQuotation" DocumentNo="Quotation_Number" Type="value" CompanyID="#Cookie.CompanyID#" LocationID="#COOKIE.LOCATION_ID#">--->
      <cfquery name="qSelectPR" datasource="#REQUEST.DSN#">
        SELECT PReq_Code FROM TACCRFQ_HEADER WHERE RFQ_Code = '#selRFQ#'
      </cfquery>
      <cfloop index="tmpPRCode" list="#qSelectPR.PReq_Code#" delimiters=",">
        <cfquery name="qUpdatePR" datasource="#REQUEST.DSN#">
          UPDATE TPPICPREQ_HEADER SET isQuotation = 1 WHERE PReq_Code = '#tmpPRCode#'
        </cfquery>
      </cfloop>
      <cfif txtSPCode eq "" and TXTMKT neq "">
        <cfquery name="qCekEmpId" datasource="#REQUEST.DSN#">
              select * from THRMEmpPersonalData
              where first_name+' '+middle_name+' '+last_name = '#txtMKT#'
          </cfquery>
      </cfif>

      <cfquery name="qNewSOL" datasource="#iif(isDefined('DSN'),'DSN','Attributes.DSN')#">
        Insert INTO TAccQuotation_Header (
          Quotation_Number,
          Quotation_Date,
          Account_ID,
          Account_Address,
          Account_Contact,
          <cfif menu eq "sales">Emp_ID,</cfif>
          <cfif menu neq "Sales">Quotation_Number_Eksternal,</cfif>
          Approval_Status,
          Quotation_Status,
          currency_id,
          Due_date,
          Quotation_Category,
          Expired,
          <!---Angries--->
          <cfif isdefined("txtNotes")>
            quotation_notes,
          </cfif>
            <cfif isdefined("txtTOPNotes")>
                Quotation_Notes,
            </cfif>
          <cfif isdefined("cboTerms") and #cboTerms# neq "">
            terms,
          </cfif>
          <cfif isdefined("qGetTypePaymentTerm") and #qGetTypePaymentTerm.recordcount#>
            Payment_Type,
          </cfif>
          <!---end--->
          quotation_type,
          company_id,
          remarks,
          WH_ID,
          QOType,
          rfq_code
          <cfif isdefined("SELPRO")>,project_code</cfif><!--- jika dia berasal dari project --->,

          Tax_Currency_ID,
          Tax_Amount,
          Base_TaxAmount,
          <cfif cookie.currencyid2 neq 0>
            base_taxamount2,
            base_amount2,
          </cfif>
          Amount,
          Base_Amount
          <!---CurrencyRateList,<!---dipindah ke table taccdocumentrate--->
          Tax_CurrencyRateList--->
          <cfif isDefined("txtQtype") AND txtQtype EQ 0>
            , Tax_Code
          </cfif>
          <cfif isDefined("SourceDoc")>
            ,SourceDoc
            ,Original_Doc
            ,Revision
            ,Revise_Reason
      ,isLead
      ,Lead_Number
          </cfif>
            <cfif Menu EQ "Sales">
              ,ShipToAddress_ID
              ,DeliveryTerm
              ,SalesQuotation_Type
              ,Sales_Type
              ,AccountSalesPerson
              ,Division_ID
              ,Bank_ID
              ,BP_Number
              <!---mik2 21/11/14 Bank GuaranteeType--->
              ,GuaranteeType
              <!---end of bank guaranteetype--->
              ,Sales_Code
              ,paymentterm_code
              ,paymentterm_code_default
        ,warranty
      </cfif>
          ,Created_By
            ,Created_Date
            <cfif Menu EQ "Sales">
            ,Vendor_ID
            <!--- ,Promised_Date --->
            ,PartialShipment
            ,VendorShipAddress
      ,Sales_Area
      ,deliveryTime
      ,AccountEnginering
            <!--- nath. 20150119. tambah AE2 & AE 3 --->
            ,AccountEnginering2
            ,AccountEnginering3
            <!--- end nath --->
      ,buffer_value
      ,totalamount_buffer
      <!--- ,Reason_Rate --->
            </cfif>
        )
        VALUES (
          '#Quotation_Number#',
          #CreateODBCDate(txtTgl)#,
          '#txtCustCode#',
          '#qaddress.account_address1#',
          '#txtCPCode#',
          <cfif menu eq "sales"><cfif txtSPCode eq "" and TXTMKT neq "">'#qCekEmpId.Emp_Id#',<cfelse>'#txtSPCode#',</cfif></cfif>
          <cfif menu neq "Sales">'#txtLamp#',</cfif>
           '0', <!--- Approval Status  0 = new,1 = checked ,2 = awaiting,3 = revised, 4 = rejected, 5 = approved--->
          <cfif txtconfirm eq 'YES'>2<cfelse>1</cfif>,  <!--- SO_Status  1 = new, 2 = open, 3 = close --->
          '#selCurrency#',
          <cfif isDefined("txtDueDate") and txtDueDate neq "">#CreateODBCDate(txtDueDate)#<cfelse>null</cfif>,
          '#Cbotype#',
          <cfif isDefined("chkExp")>1<cfelse>0</cfif>,
          <!---Angries--->
          <cfif isdefined("txtNotes")>
          '#txtNotes#',
          </cfif>
            <cfif isdefined("txtTOPNotes")>
                '#txtTOPNotes#',
            </cfif>
          <cfif isdefined("cboTerms") and #cboTerms# neq "">
          '#cboTerms#',
          </cfif>
          <cfif isdefined("qGetTypePaymentTerm") and #qGetTypePaymentTerm.recordcount#>
          '#qGetTypePaymentTerm.Top_type#',
          </cfif>
          <!---end--->
          <cfif menu eq "sales">
            'Sales',
          <cfelse>
            'Purchase',
          </cfif>
          '#companyid#',
          '#txtCatatan#',
          #COOKIE.Location_ID#,
          <cfif isDefined("txtQtype")>#txtQtype#<cfelse>''</cfif>,
          <cfif isdefined("selRFQ")>'#selRFQ#'<cfelse>''</cfif>
          <cfif isdefined("SELPRO")>,'#selPro#'</cfif><!--- jika dia berasal dari project --->
          ,'#selTaxCurrency#'
          ,#txtTotTaxConv#
          ,#txtTotTaxConv_Base#
          <cfif cookie.currencyid2 neq 0>
            ,'#txtTotTaxConv_Base2#'
            ,'#precisionevaluate(BaseInvoiceAmount2)#'
          </cfif>
          ,#InvoiceAmount#
          ,#BaseInvoiceAmount#
          <!---,'#CurrencyRateList#'
          ,'#TaxRateList#'--->
          <cfif isDefined("txtQtype") AND txtQtype EQ 0>
            ,'#listGetAt(ddlTaxIncluded, 1, "|")#'
          </cfif>
          <cfif isDefined("SourceDoc")>
            ,'#SourceDoc#'
            ,'#Original_Doc#'
            ,#Revision#
            ,'#txtReviseReason#'
      ,'#ORIisLead#'
      ,'#ORILead_Number#'
          </cfif>
            <cfif Menu EQ "Sales">
              ,#detail_id#
              ,'#selDeliveryterm#'
              ,'#selQuotationType#'
              ,'#selSOType#'
              ,'#txtAccSPCode#'
              ,#ListFirst(selDevision,"|")#
              ,#selBank#
              ,'#selBP#'
              <!---mik2 21/11/14 Bank GuaranteeType--->
              ,'#selBankGuaranteeType#'
              <!---end of bank guaranteetype--->
              ,'#selSalesCode#'
              ,'#cboTermsNew#'
              ,'#hdnTermsNew#'
        ,'#txtWarranty#'
            </cfif>
            ,#Cookie.CKSATRIADEVID#
            ,#NOW()#
            <cfif Menu EQ "Sales">
            ,#selVendor#
            <!--- ,#CreateODBCDate(txtPDate)# --->
            ,<cfif isDefined("chkPartialShip") AND chkPartialShip EQ "1">1<cfelse>0</cfif>
            ,#VAL(Vdetail_id)#
      ,'#selSalesArea#'
      ,'#txtDeliveryTime#'
      ,'#txtAESPCode#'
            <!--- nath. 20150119. tambah AE2 & AE 3 --->
            ,'#txtAESPCode2#'
            ,'#txtAESPCode3#'
            <!--- end nath --->
      ,<cfif isdefined ("rdobuff")>'#rdobuff#'<cfelse>NULL</cfif>
      ,'#INT(VAL(Replace(txtTotAmountBuffer,",","","ALL")))#'
      <!--- ,'#txtReasonRate#' --->
            </cfif>
        );

    <!---UPDATE TACCRFQ_HEADER SET EXPIRED = 1 WHERE RFQ_Code = '#selRFQ#';--->
    <cfif isDefined("SourceDoc")>
    UPDATE TAccQuotation_Header SET EXPIRED = 1 WHERE Quotation_Number = '#SourceDoc#';
    </cfif>

    <cfif isDefined("SourceDoc")>
      <cfquery name="qUpdateLead" datasource="#REQUEST.DSN#">
        UPDATE TAccQuotation_Header SET Lead_Number = '#Quotation_Number#' WHERE Lead_Number = '#SourceDoc#'
      </cfquery>
    </cfif>
    </cfquery>
    <cfquery name="qNewDocRate" datasource="#iif(isDefined('DSN'),'DSN','Attributes.DSN')#">
          insert into taccdocumentrate
          (
          document_number,
          currencyratelist,
          tax_currencyratelist
          <cfif cookie.currencyid2 neq 0><!---bila dualbase--->
            ,currencyratelist2,
            tax_currencyratelist2
          </cfif>
          )
          values
          (
          '#quotation_number#',
          '#CurrencyRateList#',
          '#taxratelist#'
          <cfif cookie.currencyid2 neq 0>
          ,'#CurrencyRateList2#',
          '#taxratelist2#'
          </cfif>
          )
      </cfquery>
    <cfelseif Task eq "Edit">
    <cfquery name="qUpdateSOL" datasource="#iif(isDefined('DSN'),'DSN','Attributes.DSN')#">
      Update TAccQuotation_Header Set
        Quotation_Date = #CreateODBCDate(txtTgl)#,
        Account_ID = '#txtCustCode#',
        Account_Address = '#qaddress.account_address1#'<!--- '#trim(replace(txtCustAddress,"#chr(13)##chr(10)#",";","all"))#' --->,
        Account_Contact = '#txtCPCode#',
        <cfif menu eq "Sales">
        Emp_ID = '#txtSPCode#',
        AccountEnginering = '#txtAESPCode#',
                AccountEnginering2 = '#txtAESPCode2#',
                AccountEnginering3 = '#txtAESPCode3#',
        deliveryTime = '#txtdeliveryTime#',
        </cfif>
        <cfif menu eq "Purchase">Quotation_Number_Eksternal = '#txtLamp#',</cfif>
        <!---Angries--->
            <cfif isdefined("txtNotes")>
          Quotation_Notes='#txtNotes#',
            </cfif>
                <cfif isdefined("txtTOPNotes")>
          Quotation_Notes='#txtTOPNotes#',
            </cfif>
        <cfif isdefined("cboTerms") and #cboTerms# neq "">
          terms='#cboTerms#',
        </cfif>
        <cfif isdefined("qGetTypePaymentTerm") and #qGetTypePaymentTerm.recordcount#>
          Payment_Type='#qGetTypePaymentTerm.Top_type#',
        </cfif>
        <!---end--->
        Approval_Status = '0',  <!--- Approval Status  0 = new,1 = checked ,2 = awaiting,3 = revised, 4 = rejected, 5 = approved--->
        Quotation_Status = <cfif txtconfirm eq 'YES'>2<cfelse>1</cfif>, <!--- Quo_Status  1 = new, 2 = open, 3 = close --->
        Expired  = <cfif isDefined("chkExp")>1<cfelse>0</cfif>,
        Due_Date= #CreateODBCDate(txtDueDate)#,
        currency_id = '#selCurrency#',
        quotation_type = <cfif menu eq "sales">'Sales'<cfelse>'Purchase'</cfif>,
        remarks = '#txtCatatan#',
        WH_ID = #COOKIE.Location_ID#,
        QOType = <cfif isDefined("txtQtype")>#txtQtype#<cfelse>''</cfif>,
        Tax_Currency_ID   = '#selTaxCurrency#',
        Tax_Amount      = #txtTotTaxConv#,
        Base_TaxAmount    = #txtTotTaxConv_Base#,

        <cfif cookie.currencyid2 neq 0>
          base_taxamount2 = '#txtTotTaxConv_Base2#',
          base_amount2 = '#precisionevaluate(baseinvoiceamount2)#',
        </cfif>

        Amount        = #InvoiceAmount#,
        Base_Amount     = #BaseInvoiceAmount#

        <!---CurrencyRateList = '#CurrencyRateList#',
        Tax_CurrencyRateList = '#TaxRateList#'--->
        <cfif isDefined("txtQtype") AND txtQtype EQ 0>
          , Tax_Code = '#listGetAt(ddlTaxIncluded, 1, "|")#'
        <cfelse>
          , Tax_Code = null
        </cfif>
          <cfif Menu EQ "Sales">
              ,ShipToAddress_ID = #detail_id#
              ,DeliveryTerm = '#selDeliveryterm#'
              ,SalesQuotation_Type = '#selQuotationType#'
              ,Sales_Type = '#selSOType#'
              ,AccountSalesPerson = '#txtAccSPCode#'
              <!---,Division_ID = #ListFirst(selDevision,"|")#--->
              ,Bank_ID = #selBank#
              ,BP_Number = '#selBP#'
              <!---mik2 21/11/14 Bank GuaranteeType--->
              ,GuaranteeType = '#selBankGuaranteeType#'
              <!---end of bank guaranteetype--->
             <!--- ,Sales_Code = '#selSalesCode#'--->
              ,paymentterm_code = '#cboTermsNew#'
        ,warranty = '#txtwarranty#'
            </cfif>
            <cfif isDefined("txtReviseReason")>
              ,Revise_Reason = '#txtReviseReason#'
      </cfif>
          ,Updated_By = #Cookie.CKSATRIADEVID#
            ,Updated_Date = #NOW()#
            <cfif Menu EQ "Sales">
            ,Vendor_ID = #selVendor#
            <!--- ,Promised_Date = #CreateODBCDate(txtPDate)# --->
            ,PartialShipment = <cfif isDefined("chkPartialShip") AND chkPartialShip EQ "1">1<cfelse>0</cfif>
            ,VendorShipAddress = #VAL(Vdetail_id)#
      ,Sales_Area = '#selSalesArea#'
      <cfif isdefined("rdobuff")>,buffer_value = '#rdoBuff#'</cfif>
      ,totalamount_buffer = '#INT(VAL(Replace(txtTotAmountBuffer,",","","ALL")))#'
            </cfif>
      <!--- ,reason_rate = '#txtReasonRate#' --->
      Where Quotation_Number = '#Quotation_Number#'
    </cfquery>
    <cfquery name="qUpdateDocRate" datasource="#iif(isDefined('DSN'),'DSN','Attributes.DSN')#">
      update taccdocumentrate
        set currencyratelist='#currencyratelist#'
        ,tax_currencyratelist='#taxratelist#'
        <cfif cookie.currencyid2 neq 0>
          ,currencyratelist2='#currencyratelist2#'
          ,tax_currencyratelist2='#taxratelist2#'
        </cfif>
      where document_number='#Quotation_Number#'
    </cfquery>
    <cfquery name="qDelSOL_Detail_Lama" datasource="#iif(isDefined('DSN'),'DSN','Attributes.DSN')#">
      Delete
      From  TAccQuotation_Detail
      Where   Quotation_Number = '#Quotation_Number#'
    </cfquery>
    </cfif>
    <cfquery datasource="#request.dsn#">
      delete from taccdocrelation where destdoc = '#Quotation_Number#'
  </cfquery>
    <cfset prevorder=0>
    <cfset setorder=0>
    <cfloop index="i" from="1" to="#rowCount#">
      <cfif isDefined("txtpartno#i#")>
        <cfif evaluate("hdndorder_#i#") neq prevorder>
          <cfset setorder=setorder+1>
          <cfset prevorder=evaluate("hdndorder_#i#")>
        </cfif>
        <cfset TotalPrice = #replace(evaluate("txtAmount"&i),",","","ALL")#>
        <!---<cfset TotalPriceBase = TotalPrice * #val(replace(evaluate("txtCurr_#selCurrency#"),",","","ALL"))#>--->
<!--- START UPDATE BY BEDU --->
        <cfif selcurrency eq cookie.currencyid>
          <cfset TotalPriceBase = PrecisionEvaluate(TotalPrice)>
          <cfelse>
          <cfset TotalPriceBase = PrecisionEvaluate(TotalPrice * replace(evaluate("txtCurr_#selCurrency#"),",","","ALL"))>
        </cfif>
        <cfset TotalPriceBase = INT(TotalPriceBase)>
<!--- END UPDATE BY BEDU --->
<!--- START UPDATE BY BEDU --->
<cfset Price = #replace(evaluate("txtConvertedUnitPrice"&i),",","","ALL")#>
        <cfif selcurrency eq cookie.currencyid>
          <cfset PriceBase = #PrecisionEvaluate(price)#>
          <cfelse>
          <cfset PriceBase = #PrecisionEvaluate(price * replace(evaluate("txtCurr_#selCurrency#"),",","","ALL"))#>
        </cfif>
        <cfset PriceBase = INT(PriceBase)>
<!--- END UPDATE BY BEDU --->
        <cfif cookie.currencyid2 neq 0>
          <!---untuk dualbase--->
          <!---
      <cfset TotalPriceBase2 = TotalPrice * #val(replace(evaluate("txtCurr2_#selCurrency#"),",","","ALL"))#>
          <cfset priceBase2= price * #val(replace(evaluate("txtCurr2_#selCurrency#"),",","","ALL"))#>
      --->
          <!--- <cfset TotalPriceBase2 = #PrecisionEvaluate(TotalPrice * #val(replace(evaluate("txtCurr2_#selCurrency#"),",","","ALL"))#)#> --->
          <cfset TotalPriceBase2 = #PrecisionEvaluate(TotalPrice * replace(evaluate("txtCurr2_#selCurrency#"),",","","ALL"))#>
          <!--- <cfset priceBase2= #PrecisionEvaluate(price * #val(replace(evaluate("txtCurr2_#selCurrency#"),",","","ALL"))#)#> --->
          <cfset priceBase2= #PrecisionEvaluate(price * replace(evaluate("txtCurr2_#selCurrency#"),",","","ALL"))#>
        </cfif>
        <cfset TaxAmount1 =  #val(replace(evaluate("txtTaxAmount1#i#"),",","","ALL"))#>
        <cfset TaxAmount2 =  #val(replace(evaluate("txtTaxAmount2#i#"),",","","ALL"))#>
        <cfset TaxAmount3 =  #val(replace(evaluate("txtTaxAmount3#i#"),",","","ALL"))#>
        <cfset TaxAmount4 =  #val(replace(evaluate("txtTaxAmount4#i#"),",","","ALL"))#>
        <cfif #evaluate('txtUnitId_#i#')#  eq "" or #evaluate('txtUnitId_#i#')#  eq "undefined">
          <cfset UnitType = 0>
          <cfelse>
          <cfset UnitType = #evaluate('txtUnitId_#i#')#>
        </cfif>
        <cfif #evaluate('txtUnitId2#i#')#  eq "" or #evaluate('txtUnitId2#i#')#  eq "undefined">
          <cfset UnitType2 = 0>
          <cfelse>
          <cfset UnitType2 = #evaluate('txtUnitId2#i#')#>
        </cfif>
        <cfif #FORM['txtDimensionID_' & i]# eq "" or  #FORM['txtDimensionID_' & i]# eq "undefined"  >
          <cfset DimensionId = 0>
          <cfelse>
          <cfset DimensionId = #FORM['txtDimensionID_' & i]#>
        </cfif>
        <cfif #form["HID_generate_flag#i#"]#  eq "" or #form["HID_generate_flag#i#"]# eq "undefined">
          <cfset generateFlag = 0>
          <cfelse>
          <cfset generateFlag = #form["HID_generate_flag#i#"]#>
        </cfif>
        <cfif listfindnocase(menu,"sales")>
          <!---Remark By mik2 11/12/14<cfif isDefined("chkPartialShip") AND chkPartialShip EQ "1">--->
          <cfquery name="qGetLeadTime#i#" datasource="#iif(isDefined('DSN'),'DSN','Attributes.DSN')#">
                        SELECT AreaCode,LeadTime_Air,LeadTime_Sea FROM TAccArea
                        WHERE Company_ID = '#Cookie.CompanyID#'
                        AND AreaCode like '#evaluate('selArea_#i#')#'
                    </cfquery>
          <cfif #evaluate('selMethod_#i#')# eq 'Sea'>
            <cfif #evaluate('qGetLeadTime#i#.LeadTime_Sea')# neq ''>
              <cfset 'leadtime_#i#' = #evaluate('qGetLeadTime#i#.LeadTime_Sea')#>
              <cfelse>
              <cfset 'leadtime_#i#' = 0>
            </cfif>
            <cfelseif #evaluate('selMethod_#i#')# eq 'Air'>
            <cfif #evaluate('qGetLeadTime#i#.LeadTime_Air')# neq ''>
              <cfset 'leadtime_#i#' = #evaluate('qGetLeadTime#i#.LeadTime_Air')#>
              <cfelse>
              <cfset 'leadtime_#i#' = 0>
            </cfif>
            <cfelse>
            <cfset 'leadtime_#i#' = 0>
          </cfif>
          <cfif #evaluate('TXTEXWORKFACTORY_#i#')# eq ''>
            <cfset 'TXTEXWORKFACTORY_#i#' = 0>
          </cfif>
          <cfset 'totaltime_#i#' = #evaluate('leadtime_#i#')#+#evaluate('TXTEXWORKFACTORY_#i#')#>
          <cfset 'txtPDate#i#' = #dateADD('d',evaluate('totaltime_#i#'),txttgl)#>
          <!---</cfif>--->
          <cfquery name="qSOL_Detail#i#" datasource="#iif(isDefined('DSN'),'DSN','Attributes.DSN')#">
          Insert Into TAccQuotation_Detail (
            Quotation_number,
            Item_Code,
            Item_Desc,
            Qty,
            UnitPrice
            ,qty2
            , Unit_Type
            , Unit_Type2

            ,Base_UnitPrice,
            <cfif cookie.currencyid2 neq 0>
              base_unitprice2,
              base_totalprice2,
            </cfif>
            Disc_percentage,

            Tax_Code1,
            Tax_Percentage1,
            Tax_Operator1,
            Tax_Amount1,

            Tax_Code2,
            Tax_Percentage2,
            Tax_Operator2,
            Tax_Amount2,

            <!--- TAX 3 & 4 --->
            Tax_Code3,
            Tax_Percentage3,
            Tax_Operator3,
            Tax_Amount3,

            Tax_Code4,
            Tax_Percentage4,
            Tax_Operator4,
            Tax_Amount4,

            TotalPrice,
            Base_TotalPrice,
            generate_flag,
            parent_item,
            parent_path,
            <!--- colour, --->


            config_level,
            config_ratio,
            config_order,
            disc_type,
            Dimension_ID,
            Disc_Value
                        ,DetailPromisedDate
                        <!---mik2 8/12/14--->
                        ,LeadTime
                        ,ExWorkFactory
                        ,LeadTimeDays
          )
          values (
             '#Quotation_Number#',
             '#evaluate('txtpartNo#i#')#',
             <cfqueryparam cfsqltype="cf_sql_varchar" value="#evaluate('txtdesc#i#')#"/>,
             #replace(evaluate('txtQty_#i#'),",","","ALL")#,
             #Price#
             , #val(replace(evaluate('txtQty2#i#'),",","","ALL"))#
             , #UnitType#
             , #UnitType2#

             ,#PriceBase#,
             <cfif cookie.currencyid2 neq 0>
               '#pricebase2#',
               '#totalpricebase2#',
             </cfif>
             '#replace(evaluate('txtDiscount#i#'),",","","ALL")#',

             '#listgetat(evaluate("seltax1"&i),1,"|")#',
             #listgetat(evaluate("seltax1"&i),2,"|")#,
             '#listgetat(evaluate("seltax1"&i),3,"|")#',
             #TaxAmount1#,

             '#listgetat(evaluate("seltax2"&i),1,"|")#',
             #listgetat(evaluate("seltax2"&i),2,"|")#,
             '#listgetat(evaluate("seltax2"&i),3,"|")#',
             #TaxAmount2#,

             <!--- TAX 3 & 4 --->
             '#listgetat(evaluate("seltax3"&i),1,"|")#',
             #listgetat(evaluate("seltax3"&i),2,"|")#,
             '#listgetat(evaluate("seltax3"&i),3,"|")#',
             #TaxAmount3#,

             '#listgetat(evaluate("seltax4"&i),1,"|")#',
             #listgetat(evaluate("seltax4"&i),2,"|")#,
             '#listgetat(evaluate("seltax4"&i),3,"|")#',
             #TaxAmount4#,

             #TotalPrice#,
             #TotalPriceBase#,
             '#form["HID_generate_flag#i#"]#',
             '#evaluate ("hid_parent_item#i#")#',
             '<cfif evaluate("form.parent_path#i#") IS "undefined">0<cfelse>#form["parent_path#i#"]#</cfif>'
            <!---  '#evaluate("chkwarna#i#")#', --->


             ,<cfif isdefined("form.hdnLevel#i#")>#val(form["hdnLevel#i#"])#<cfelse>NULL</cfif>
             ,<cfif isdefined("form.hdnRatio_#i#")>#val(form["hdnRatio_#i#"])#<cfelse>NULL</cfif>
             ,#setorder#
             ,'<cfif replace(evaluate('txtDiscType_#i#'),",","","ALL") neq "undefined">#replace(evaluate('txtDiscType_#i#'),",","","ALL")#<cfelse>0</cfif>'
             ,#DimensionId#
             ,#objDummy.cfnumval(objParam: FORM['txtDiscValue' & i])#
            <!---Remark By mik2 11/12/14<cfif isDefined("chkPartialShip") AND chkPartialShip EQ "1">--->
                          ,<cfif isDefined("txtPDate#i#")>#CreateODBCDate(Evaluate("txtPDate#i#"))#<cfelse>NULL</cfif>
                            <!---mik2 8/12/14--->
                            ,<cfif isDefined("selArea_#i#")>'#evaluate('selArea_#i#')#|#evaluate('selMethod_#i#')#'<cfelse>NULL</cfif>
                          ,<cfif isDefined("TXTEXWORKFACTORY_#i#")>#evaluate('TXTEXWORKFACTORY_#i#')#<cfelse>NULL</cfif>
                            ,<cfif isDefined("totaltime_#i#")>#evaluate('totaltime_#i#')#<cfelse>NULL</cfif>
                        <!---Remark By mik2 11/12/14<cfelseif isDefined("txtPDate")>
                          ,#CreateODBCDate(txtPDate)#
                            ,NULL
                            ,NULL
                        <cfelse>
                          ,NULL
                            ,NULL
                            ,NULL
                        </cfif>--->

          );

          SELECT @@IDENTITY AS 'Identity';
        </cfquery>
          <cfelse>
          <cfquery name="qSOL_Detail#i#" datasource="#iif(isDefined('DSN'),'DSN','Attributes.DSN')#">
          Insert Into TAccQuotation_Detail (
            Quotation_number,
            Item_Code,
            Item_Desc,
            Qty,
            UnitPrice
            ,qty2
            ,Unit_Type
            ,Unit_Type2

            ,Base_UnitPrice,
            <cfif cookie.currencyid2 neq 0>
              base_unitprice2,
              base_totalprice2,
            </cfif>
            Disc_percentage,

            Tax_Code1,
            Tax_Percentage1,
            Tax_Operator1,
            Tax_Amount1,

            Tax_Code2,
            Tax_Percentage2,
            Tax_Operator2,
            Tax_Amount2,

            <!--- TAX 3 & 4 --->
            Tax_Code3,
            Tax_Percentage3,
            Tax_Operator3,
            Tax_Amount3,

            Tax_Code4,
            Tax_Percentage4,
            Tax_Operator4,
            Tax_Amount4,

            TotalPrice,
            Base_TotalPrice,
            generate_flag,
            parent_item,
            parent_path,
            <!--- colour, --->


            config_level,
            config_ratio,
            config_order,
            disc_type,
            Dimension_ID,
            Disc_Value,
            Remark

          )
          values (
             '#Quotation_Number#',
             '#evaluate('txtpartNo#i#')#',
             <cfqueryparam cfsqltype="cf_sql_varchar" value="#evaluate('txtdesc#i#')#"/>,
             #replace(evaluate('txtQty_#i#'),",","","ALL")#,
             #Price#
             , #val(replace(evaluate('txtQty2#i#'),",","","ALL"))#
             , #UnitType#
             , #UnitType2#

             ,#PriceBase#,
             <cfif isdefined("cookie.CURRENCYID2") and cookie.CURRENCYID2 neq 0>
               '#pricebase2#',
               '#totalpricebase2#',
             </cfif>
             '#replace(evaluate('txtDiscount#i#'),",","","ALL")#',

             '#listgetat(evaluate("seltax1"&i),1,"|")#',
             #listgetat(evaluate("seltax1"&i),2,"|")#,
             '#listgetat(evaluate("seltax1"&i),3,"|")#',
             #TaxAmount1#,

             '#listgetat(evaluate("seltax2"&i),1,"|")#',
             #listgetat(evaluate("seltax2"&i),2,"|")#,
             '#listgetat(evaluate("seltax2"&i),3,"|")#',
             #TaxAmount2#,

             <!--- TAX 3 & 4 --->
             '#listgetat(evaluate("seltax3"&i),1,"|")#',
             #listgetat(evaluate("seltax3"&i),2,"|")#,
             '#listgetat(evaluate("seltax3"&i),3,"|")#',
             #TaxAmount3#,

             '#listgetat(evaluate("seltax4"&i),1,"|")#',
             #listgetat(evaluate("seltax4"&i),2,"|")#,
             '#listgetat(evaluate("seltax4"&i),3,"|")#',
             #TaxAmount4#,

             #TotalPrice#,
             #TotalPriceBase#,
             '#generateFlag#',
             '#evaluate ("hid_parent_item#i#")#',
             '<cfif evaluate("form.parent_path#i#") IS "undefined">0<cfelse>#form["parent_path#i#"]#</cfif>'
            <!---  '#evaluate("chkwarna#i#")#', --->


             ,<cfif isdefined("form.hdnLevel#i#")>#val(form["hdnLevel#i#"])#<cfelse>NULL</cfif>
             ,<cfif isdefined("form.hdnRatio_#i#")>#val(form["hdnRatio_#i#"])#<cfelse>NULL</cfif>
             ,#setorder#
             , '<cfif replace(evaluate('txtDiscType_#i#'),",","","ALL") neq "undefined">#replace(evaluate('txtDiscType_#i#'),",","","ALL")#<cfelse>0</cfif>'
             ,#DimensionId#
             ,#objDummy.cfnumval(objParam: FORM['txtDiscValue' & i])#
             , '#evaluate('txtRemark_#i#')#'
          );

          SELECT @@IDENTITY AS 'Identity';
        </cfquery>
        </cfif>
        <cfif isDefined("hdnpr#i#")>
          <cfif len(evaluate("hdnpr#i#"))>
            <cfquery  datasource="#iif(isdefined('DSN'),'DSN','Attributes.DSN')#">
                      insert into taccdocrelation (type, sourcedoc, destdoc, itemcode, company_id, itemqty,Dimension_ID)
                      values (<cfif submenu eq "purchase">'PR-QUOTATION'<cfelse>'SO-QUOTATION'</cfif>, '#evaluate("hdnpr#i#")#', '#Quotation_Number#', '#evaluate("txtPartNo"&i)#', #cookie.companyid#, #replace(evaluate('txtQty_#i#'),",","","ALL")#,#FORM['txtDimensionID_' & i]#)
                  </cfquery>
          </cfif>
        </cfif>
        <!---   <br/> #Quotation_Number# - #evaluate('txtpartNo#i#')# --->
        <!---- Begin of -- simpan data warna jika item punya warna --->
        <!---- End of -- simpan data warna jika item punya warna --->
      </cfif>
    </cfloop>
    <!--- Start Upload Item Detail CRF51014-14497 YS 20141008--->
    <cfset OBJName = "UploadItem">
    <cfif isDefined("#OBJName#RowCount") AND Evaluate("#OBJName#RowCount") NEQ 0>
      <cfset Row = Evaluate("#OBJName#RowCount")>
      <cfif selCurrency EQ cookie.currencyid>
        <cfset Rate = 1>
        <cfelse>
        <cfset Rate = Replace(Evaluate("txtCurr_#selCurrency#"),",","","ALL")>
      </cfif>
      <cfset Rate2 = Replace(Evaluate("txtCurr2_#selCurrency#"),",","","ALL")>
      <cfloop from="1" to="#Row#" index="i">
        <cfif isDefined("#OBJName#txtPartNo#i#") AND Evaluate("#OBJName#txtPartNo#i#") NEQ "">
          <cfset TotalPrice = Replace(Evaluate("#OBJName#txtAmount#i#"),",","","ALL")>
          <cfset TotalPriceBase = PrecisionEvaluate(TotalPrice * Rate)>
          <cfset Price = Replace(Evaluate("#OBJName#txtConvertedUnitPrice#i#"),",","","ALL")>
          <cfset PriceBase = PrecisionEvaluate(Price * Rate)>
          <!--- Base 2 --->
          <!--- START UPDATE BY BEDU --->
          <cfset TotalPriceBase = INT(TotalPriceBase)>
          <cfset PriceBase = INT(PriceBase)>
          <cfif Cookie.CurrencyID2 NEQ 0>
      <cfset TotalPriceBase2 = PrecisionEvaluate(TotalPrice * Rate2)>
          <cfset PriceBase2 = PrecisionEvaluate(Price * Rate2)>
      <cfset TotalPriceBase2 = INT(TotalPriceBase2)>
          <cfset PriceBase2 = INT(PriceBase2)>  
          </cfif>
          <!--- END UPDATE BY BEDU --->
          <cfset TaxAmount1 =  Replace(Evaluate("#OBJName#txtTaxAmount1#i#"),",","","ALL")>
          <cfset TaxAmount2 =  Replace(Evaluate("#OBJName#txtTaxAmount2#i#"),",","","ALL")>
          <cfset TaxAmount3 =  Replace(Evaluate("#OBJName#txtTaxAmount3#i#"),",","","ALL")>
          <cfset TaxAmount4 =  Replace(Evaluate("#OBJName#txtTaxAmount4#i#"),",","","ALL")>
          <cfset GenerateFlag = 0>
          <cfset ItemCode = Evaluate("#OBJName#txtpartNo#i#")>
          <cfset ItemName = Evaluate("#OBJName#txtdesc#i#")>
          <cfset ItemQty = Replace(Evaluate("#OBJName#txtQty_#i#"),",","","ALL")>
          <cfset DiscPerc = Replace(Evaluate("#OBJName#txtDiscount#i#"),",","","ALL")>
          <cfset DiscValue = Replace(Evaluate("#OBJName#txtDiscValue#i#"),",","","ALL")>
          <cfset Tax1 = Evaluate("#OBJName#selTax1#i#")>
          <cfset Tax2 = Evaluate("#OBJName#selTax2#i#")>
          <cfset Tax3 = Evaluate("#OBJName#selTax3#i#")>
          <cfset Tax4 = Evaluate("#OBJName#selTax4#i#")>
          <!---mik2 11/12/14--->
          <cfif isDefined("#OBJName#selArea_#i#")>
            <cfset ItemArea = Evaluate("#OBJName#selArea_#i#")>
            <cfset ItemMethod = Evaluate("#OBJName#selMethod_#i#")>
            <cfset ItemExWorkFactory = Evaluate("#OBJName#txtExworkFactory_#i#")>
            <cfquery name="qGetLeadTime#i#" datasource="#iif(isDefined('DSN'),'DSN','Attributes.DSN')#">
                        SELECT AreaCode,LeadTime_Air,LeadTime_Sea FROM TAccArea
                        WHERE Company_ID = '#Cookie.CompanyID#'
                        AND AreaCode like '#ItemArea#'
                    </cfquery>
            <cfif ItemMethod eq 'Sea'>
              <cfif #evaluate('qGetLeadTime#i#.LeadTime_Sea')# neq ''>
                <cfset 'leadtime_#i#' = #evaluate('qGetLeadTime#i#.LeadTime_Sea')#>
                <cfelse>
                <cfset 'leadtime_#i#' = 0>
              </cfif>
              <cfelseif ItemMethod eq 'Air'>
              <cfif #evaluate('qGetLeadTime#i#.LeadTime_Air')# neq ''>
                <cfset 'leadtime_#i#' = #evaluate('qGetLeadTime#i#.LeadTime_Air')#>
                <cfelse>
                <cfset 'leadtime_#i#' = 0>
              </cfif>
              <cfelse>
              <cfset 'leadtime_#i#' = 0>
            </cfif>
            <cfif TRIM(ItemExWorkFactory) EQ "">
              <cfset ItemExWorkFactory = 0>
            </cfif>
            <cfset 'totaltime_#i#' = #evaluate('leadtime_#i#')#+#ItemExWorkFactory#>
            <cfset 'txtPDate#i#' = #dateADD('d',evaluate('totaltime_#i#'),txttgl)#>
          </cfif>
          <!--- Get From Master Item --->
          <cfquery name="qGetItemInfo" datasource="#iif(isDefined('DSN'),'DSN','Attributes.DSN')#">
                  SELECT Unit_Type_ID,Dimension_ID FROM TItem INNER JOIN TItemCompany ON TItem.Item_Code = TItemCompany.Item_Code
                    WHERE Company_ID = #Cookie.CompanyID# AND TItem.Item_Code = '#ItemCode#'
                </cfquery>
          <cfset UnitType = qGetItemInfo.Unit_Type_ID>
          <cfset UnitType2 = UnitType>
          <cfset DimensionID = qGetItemInfo.Dimension_ID>
          <cfquery name="qSOL_Detail#i#" datasource="#iif(isDefined('DSN'),'DSN','Attributes.DSN')#">
                    INSERT INTO TAccQuotation_Detail (
                        Quotation_Number,Item_Code,Item_Desc,Qty,UnitPrice
                        ,Qty2,Unit_Type,Unit_Type2,Base_UnitPrice,
                        <cfif Cookie.CurrencyID2 NEQ 0>
                            Base_UnitPrice2,
                            Base_TotalPrice2,
                        </cfif>
                        Disc_Percentage,
                        Tax_Code1,Tax_Percentage1,Tax_Operator1,Tax_Amount1,
                        Tax_Code2,Tax_Percentage2,Tax_Operator2,Tax_Amount2,
                        Tax_Code3,Tax_Percentage3,Tax_Operator3,Tax_Amount3,
                        Tax_Code4,Tax_Percentage4,Tax_Operator4,Tax_Amount4,
                        TotalPrice, Base_TotalPrice,
                        Generate_Flag,Parent_Item,Parent_Path,Config_Level,Config_Ratio,Config_Order,
                        Disc_Type,Dimension_ID,Disc_Value
                        <cfif NOT ListFindNoCase(Menu,"Sales")>
                            ,Remark
                        </cfif>
                        ,DetailPromisedDate
                         <!---mik2 8/12/14--->
                        ,LeadTime
                        ,ExWorkFactory
                        ,LeadTimeDays
                    )
                    VALUES (
                        '#Quotation_Number#','#ItemCode#',<cfqueryparam cfsqltype="cf_sql_varchar" value="#ItemName#"/>,
                        #ItemQty#,#Price#
                        ,#ItemQty#,#UnitType#,#UnitType2#,#PriceBase#,
                        <cfif Cookie.CurrencyID2 NEQ 0>
                            '#precisionevaluate(PriceBase2)#',
                            '#precisionevaluate(TotalPriceBase2)#',
                        </cfif>
                        '#precisionevaluate(DiscPerc)#',
                        '#ListGetAt(Tax1,1,"|")#',#ListGetAt(Tax1,2,"|")#,'#ListGetAt(Tax1,3,"|")#',#TaxAmount1#,
                        '#ListGetAt(Tax2,1,"|")#',#ListGetAt(Tax2,2,"|")#,'#ListGetAt(Tax2,3,"|")#',#TaxAmount2#,
                        '#ListGetAt(Tax3,1,"|")#',#ListGetAt(Tax3,2,"|")#,'#ListGetAt(Tax3,3,"|")#',#TaxAmount3#,
                        '#ListGetAt(Tax4,1,"|")#',#ListGetAt(Tax4,2,"|")#,'#ListGetAt(Tax4,3,"|")#',#TaxAmount4#,
                        #TotalPrice#, #TotalPriceBase#,
                        '#GenerateFlag#','0','0',NULL,NULL,0
                        ,'0',#DimensionID#,#DiscValue#
                        <cfif NOT ListFindNoCase(Menu,"Sales")>
                            ,'#Evaluate("#OBJName#txtRemark_#i#")#'
                        </cfif>
                        <!---Remark by mik2 11/12/14
            <cfif isDefined("txtPDate")>
                          ,#CreateODBCDate(txtPDate)#
                        <cfelse>
                          ,NULL
                        </cfif>--->
                        <!---Remark By mik2 11/12/14<cfif isDefined("chkPartialShip") AND chkPartialShip EQ "1">--->
                          ,<cfif isDefined("txtPDate#i#")>#CreateODBCDate(Evaluate("txtPDate#i#"))#<cfelse>NULL</cfif>
                            <!---mik2 8/12/14--->
                            ,<cfif isDefined("ItemArea")>'#ItemArea#|#ItemMethod#'<cfelse>NULL</cfif>
                          ,<cfif isDefined("ItemExWorkFactory")>#trim(ItemExWorkFactory)#<cfelse>NULL</cfif>
                            ,<cfif isDefined("totaltime_#i#")>#REPLACE(evaluate('totaltime_#i#')," ","","ALL")#<cfelse>NULL</cfif>
                        <!---Remark By mik2 11/12/14<cfelseif isDefined("txtPDate")>
                          ,#CreateODBCDate(txtPDate)#
                            ,NULL
                            ,NULL
                        <cfelse>
                          ,NULL
                            ,NULL
                            ,NULL
                        </cfif>--->
                    );

                    SELECT @@IDENTITY AS 'Identity';
                </cfquery>
        </cfif>
      </cfloop>
    </cfif>
    <!--- End Upload Item Detail CRF51014-14497 --->
    <!--- Update Lead Time Header --->
    <cfquery name="qUpdateLeadTimeHeader" datasource="#iif(isDefined('DSN'),'DSN','Attributes.DSN')#">
      UPDATE TAccQuotation_Header SET LeadTime = (SELECT MAX(LeadTimeDays) FROM TAccQuotation_Detail WHERE Quotation_Number = TAccQuotation_Header.Quotation_Number)
        WHERE Quotation_Number = '#Quotation_Number#'
    </cfquery>
    <cfif txtconfirm eq 'YES'>
      <CF_DO_V30_REQUESTAPPROVAL COMPANY_ID="#COOKIE.COMPANYID#" dsn="#iif(isDefined('DSN'),'DSN','Attributes.DSN')#" VST_IDX="#VST_IDX#" RequestApproval_Name="GSSOL" ReqApproval_ID="#Quotation_Number#" qty="0" amount="#BaseInvoiceAmount#" lastStatus="LASTSTATUS" IsUpdatable="ok" AutoApproval="autocreate_var"/>
    </cfif>
    <!---  @<strong>testing dualbase</strong><br>
  <cfquery name="tes" datasource="#iif(isdefined('DSN'),'DSN','Attributes.DSN')#">
      select * from taccdocumentrate where document_number='#quotation_number#'
  </cfquery>
  <cfquery name="tes2" datasource="#iif(isdefined('DSN'),'DSN','Attributes.DSN')#">
      select * from taccquotation_detail where quotation_number='#quotation_number#'
  </cfquery>
  <cfquery name="tes3" datasource="#iif(isdefined('DSN'),'DSN','Attributes.DSN')#">
      select * from taccquotation_header where quotation_number='#quotation_number#'
  </cfquery>
  <cfdump var="#tes#">
  <cfdump var="#tes2#">
  <cfdump var="#tes3#">
  <cfabort>
--->
    <!---<cfabort>--->
    <!---<cfabort>--->
    <cfif REMOTE_ADDR eq '192.168.4.160'>
      <!---<cfabort>--->
    </cfif>
    <!--- <cfabort> --->
  </cftransaction>
</cfoutput>
<!--- wx :: utk autoapproved --->
<cfif txtconfirm eq 'YES' and autocreate_var eq 1>
  <cfquery name="qEmpData" datasource="#iif(isDefined('DSN'),'DSN','Attributes.DSN')#">
    select
      THRMEmpPersonalData.Emp_ID,
      ThrmEmpPosition.Position_ID
    from THRMEmpPersonalData
    inner join ThrmEmpPosition on ThrmEmpPosition.Emp_ID = THRMEmpPersonalData.Emp_ID
    inner join ThrmPosition on ThrmEmpPosition.position_id = ThrmPosition.position_id
      WHERE ThrmEmpPosition.company_id = #cookie.companyid#
            and THRMEmpPersonalData.User_ID = '#COOKIE.CKSATRIADEVID#'
      AND THRMPosition.Position_Flag = 3
    </cfquery>
  <cfset lstEmpPosition = valueList(qEmpData.Position_ID)>
  <cfquery name="qGetApprovalData" datasource="#iif(isDefined('DSN'),'DSN','Attributes.DSN')#">
        SELECT THRMApprovedBy.ApprovedBy_ID,
          isNull(CONVERT(VARCHAR,Approved_By),THRMApprovedBy.LstApprovedBy) Approved_By,  <!--- posisi Approver --->
          THRMApprovedBy.Is_Required,
          THRMApprovedBy.Flag_Turn
        FROM    THRMApprovedBy
        WHERE   ReqApproval_ID = '#Quotation_Number#'
        order by SettingApproval_StepData, ApprovedBy_Id asc
    </cfquery>
  <cfset form.hdnDO_2009_Approval = lstEmpPosition>
  <cfset form.hdnApprovedById = "">
  <cfloop query="qGetApprovalData">
    <cfset temporary.HasAccess = "false"/>
    <cfif qGetApprovalData.Flag_Turn eq 1>
      <cfloop list="#qGetApprovalData.Approved_By#" index="pos">
        <cfif ListFind(lstEmpPosition,pos) gt 0>
          <cfset temporary.HasAccess = "true"/>
          <cfset form.hdnApproveBy = pos>
        </cfif>
      </cfloop>
    </cfif>
    <cfif temporary.HasAccess eq "true">
      <cfset form.hdnApprovedById = ListAppend(form.hdnApprovedById,qGetApprovalData.ApprovedBy_id,",")>
      <cfset "FORM.cboStatus#qGetApprovalData.ApprovedBy_id#" = 3>
      <cfset "FORM.txtReason#qGetApprovalData.ApprovedBy_id#" = "Automatic Approved By System">
    </cfif>
  </cfloop>
  <cfif form.hdnApprovedById neq "">
    <cfset noreload = 1>
    <cfinclude template="#Application.stApp.CFWeb_Path[1]#/eaccounting/sales/quotation_inbox/queries/qeditstatus.cfm">
  </cfif>
</cfif>
<cfoutput>
  <script language="JavaScript">
  <cfif TASK neq "edit">
    alert("#DO_VAR['QuotationNo']# : #Quotation_Number#")
  <cfelse>
    alert("#DO_VAR['eHRMUpdatedSuccess']#")
  </cfif>
  <cfif submenu eq "sales">
    location.href = '#Application.stApp.Web_Path[VST_IDX]#/#Application.stApp.Home_Url[VST_IDX]#/index.cfm?FID=ERSTD08033&fuid=ERSTD0803301&HelpCategory_ID=eAccSales&menu=1&submenu=#menu#&cbotype=#cbotype#&Help_ID=QuotationLetter&refresh=#URLEncodedFormat(now())#'
  <cfelse>
    location.href = '#Application.stApp.Web_Path[VST_IDX]#/#Application.stApp.Home_Url[VST_IDX]#/index.cfm?FID=ERSTD07986&fuid=ERSTD0798601&HelpCategory_ID=eAccSales&menu=1&submenu=#menu#&cbotype=#cbotype#&Help_ID=QuotationLetter&refresh=#URLEncodedFormat(now())#'
  </cfif>

</script>
</cfoutput>
