using System;
using System.Collections;
using Skrypton.RuntimeSupport;
namespace TranslatedProgram
{
    public sealed class Runner : RunnerBaseT<EnvironmentReferences, GlobalReferences>
    {
        private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
        public Runner(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer) : base(compatLayer)
        {
            _ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
        }
        protected override GlobalReferences CreateGlobalReferences(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env) => new GlobalReferences(compatLayer, env);
        protected override void Go(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env, GlobalReferences globalReferences)
        {
            var _env = env ?? throw new ArgumentNullException(nameof(env));
            var _outer = globalReferences ?? throw new ArgumentNullException(nameof(globalReferences));

            _outer.InterfaceVersion = (Int16)1;

            //nasty globals
            _outer.g_iNumberOfCalendarsRendered = (Int16)0;
            _outer.bFormRendered = false;

            _outer.bProdHasAvail = false;

            //public methods

            //internal methods

            // ==========================================================================================================
            // These functions are all about assigning unit selections when the booking process is hooked into from an
            // external site (eg. VisitBritain).
            //
            // The other site will have requested availability data through the webservice and the order in which
            // the options appear there may vary from the order that they're returned from the availability object's
            // queries.
            //
            // These functions are intended to pick up on selections in the querystring and match them back up to
            // the ReqNo entries.
            //
            // Selections passed in are given as parameters of the form
            //  URslt1=12345,1,1
            // where the comma-separated values are UnitKey, number of adults, number of children.
            // The "1" in "URslt1" is expected to match up with ReqNo 1 when passed through, but the problem is that
            // this often isn't the case.
            // ==========================================================================================================

        }
    }
    public sealed class GlobalReferences : GlobalReferencesBaseT<EnvironmentReferences>
    {
        private readonly IProvideVBScriptCompatFunctionalityToIndividualRequests _;
        private readonly GlobalReferences _outer;
        private readonly EnvironmentReferences _env;
        public GlobalReferences(IProvideVBScriptCompatFunctionalityToIndividualRequests compatLayer, EnvironmentReferences env) : base(compatLayer, env)
        {
            _ = compatLayer ?? throw new ArgumentNullException(nameof(compatLayer));
            _env = env ?? throw new ArgumentNullException(nameof(env));
            _outer = this;
            Page = null;
            Request = null;
            Context = null;
            Server = null;
            DMS = null;
            InterfaceVersion = null;
            BOOKING_Local = (Int16)0;
            BOOKING_Eviivo = (Int16)1;
            BOOKING_Redirect = (Int16)2;
            BOOKING_PollingRedirect = (Int16)3;
            IsExternalBooking = null;
            strExtBookUrl = null;
            strProductEstateID = null;
            bFormRendered = null;
            IsVBPollingEnabled = null;
            bRenderAsCalendar = null;
            g_iNumberOfCalendarsRendered = null;
            bProdHasAvail = null;
        }
        internal object Page { get; set; }
        internal object Request { get; set; }
        internal object Context { get; set; }
        internal object Server { get; set; }
        internal object DMS { get; set; }
        internal object InterfaceVersion { get; set; }
        internal object BOOKING_Local { get; set; }
        internal object BOOKING_Eviivo { get; set; }
        internal object BOOKING_Redirect { get; set; }
        internal object BOOKING_PollingRedirect { get; set; }
        internal object IsExternalBooking { get; set; }
        internal object strExtBookUrl { get; set; }
        internal object strProductEstateID { get; set; }
        internal object bFormRendered { get; set; }
        internal object IsVBPollingEnabled { get; set; }
        internal object bRenderAsCalendar { get; set; }
        internal object g_iNumberOfCalendarsRendered { get; set; }
        internal object bProdHasAvail { get; set; }
        public object GetProdHasAvail()
        {
            return _.VAL(_outer.bProdHasAvail);
        }
        // ====================================================================================================
        // RENDER: Availability Calendar (supports local availability only!)
        // - Note: This doesn't actually perform any data access, all of the content required is passed
        //   through in POST data from the availability calendar on the previous page (Product Detail)
        // ====================================================================================================
        public object BookingUI_StayMain_AvailCal(object pO, object objRenderSettings)
        {
            object BookingUI_StayMain_AvailCal_retVal = null;
            object objBookingRequirement = null;
            object intBookingType = null;
            object iStayNum = null;
            object iThisReqmnt = null;
            object iUnitQty = null;
            object iUnitMinOccupancy = null;
            object iUnitMaxCapacity = null;
            object iUnitKey = null;
            object iLinkedUnitKey = null;
            object strUnitName = null;
            object strAvailClassId = null;
            object Item = null;
            object strTemp = null;
            object i = null;
            object dStart = null;
            object iNights = null;
            object intProdKey = null;
            // Expect selections as set of form values:
            //  "unit_prodkey", "minoccu_prodkey", "maxcap_prodkey", "name_prodkey", "availclass_prodkey"
            //
            // If linked units are referenced, the first value will be:
            //  "unit_prodkey_linkprodkey"

            // 2011-08-09 DWR: Get populated read-only Booking Requirement data From GetSharedObject, then translate into a local copy we can edit
            // (since some methods in here try to mess about with properties on it)
            objBookingRequirement = _.OBJ(_.CALLm2v1(this, _outer.Page, "Functions", "GetSharedObject", "BookingRequirement"));
            objBookingRequirement = _.OBJ(_.CALLm1argp(this, _outer, "GetEditableBookingRequirement", _.ARGS.Ref(objBookingRequirement, v => { objBookingRequirement = v; })));

            dStart = _.VAL(_.CALLm1v0(this, objBookingRequirement, "VisitDate"));
            iNights = _.VAL(_.CALLm1v0(this, objBookingRequirement, "Nights"));
            intProdKey = _.VAL(_.CALLm1v0(this, objBookingRequirement, "Product"));

            // Open form and prepare to wrap content in "staySelection" container
            if (_.IF(_outer.IsExternalBooking))
            {
                intBookingType = _.VAL(_outer.BOOKING_Redirect);
            }
            else
            {
                intBookingType = _.VAL(_outer.BOOKING_Local);
            }

            _.CALLm1argp(this, _outer, "RenderBookingInfoForm", _.ARGS.Ref(pO, v2 => { pO = v2; }).Ref(intProdKey, v3 => { intProdKey = v3; }).Ref(objRenderSettings, v4 => { objRenderSettings = v4; }).Ref(intBookingType, v5 => { intBookingType = v5; }).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null));

            _.CALLm1v1(this, pO, "Write", "<div class=\"staySelection\">");

            // Try to pull requirement info from Request
            iStayNum = (Int16)1;
            iThisReqmnt = (Int16)0;
            var enumerationContent = _.ENUMERABLE(_.CALLm1v0(this, _outer.Request, "Form")).GetEnumerator();
            while (true)
            {
                if (!enumerationContent.MoveNext())
                    break;
                Item = enumerationContent.Current;
                //## Loop through only units
                if (_.IF(_.EQ(_.NullableSTR(_.LEFT(Item, (Int16)5)), "unit_")))
                {

                    strTemp = _.VAL(_.RIGHT(Item, _.SUBT(_.LEN(Item), (Int16)5)));
                    if (_.IF(_.GT(_.NullableNUM(_.INSTR(strTemp, "_")), (Int16)0)))
                    {
                        // Linked unit
                        iUnitKey = _.CLNG(_.RIGHT(strTemp, _.SUBT(_.LEN(strTemp), _.INSTR(strTemp, "_"))));
                        iLinkedUnitKey = _.CLNG(_.LEFT(strTemp, _.SUBT(_.INSTR(strTemp, "_"), (Int16)1)));
                    }
                    else
                    {
                        iUnitKey = _.CLNG(strTemp);
                        iLinkedUnitKey = (Int16)0;
                    }

                    iUnitQty = _.CLNG(_.CONCAT("0", _.CALLm0argp(this, _outer.Request, _.ARGS.Ref(Item, v6 => { Item = v6; }))));
                    iUnitMinOccupancy = _.CLNG(_.CONCAT("0", _.CALLm0argp(this, _outer.Request, _.ARGS.Val(_.CONCAT("minoccu_", strTemp)))));
                    iUnitMaxCapacity = _.CLNG(_.CONCAT("0", _.CALLm0argp(this, _outer.Request, _.ARGS.Val(_.CONCAT("maxcap_", strTemp)))));

                    strUnitName = _.VAL(_.CALLm0argp(this, _outer.Request, _.ARGS.Val(_.CONCAT("name_", strTemp))));
                    strAvailClassId = _.VAL(_.CALLm0argp(this, _outer.Request, _.ARGS.Val(_.CONCAT("availclass_", strTemp))));
                    if (_.IF(_.GT(_.NullableNUM(iUnitQty), (Int16)0)))
                    {
                        var loopEnd = _.NUM(iUnitQty);
                        var loopStart = _.NUM((Int16)1, loopEnd);
                        if (_.StrictLTE(loopStart, loopEnd))
                        {
                            for (i = loopStart; _.StrictLTE(i, loopEnd); i = _.ADD(i, (Int16)1))
                            {
                                iThisReqmnt = _.ADD(iThisReqmnt, (Int16)1);
                                _.CALLm1argp(this, _outer, "BookingUI_RenderNewReq_AvailCal", _.ARGS.Ref(intBookingType, v7 => { intBookingType = v7; }).Ref(iUnitKey, v8 => { iUnitKey = v8; }).Ref(strUnitName, v9 => { strUnitName = v9; }).Ref(iUnitMinOccupancy, v10 => { iUnitMinOccupancy = v10; }).Ref(iUnitMaxCapacity, v11 => { iUnitMaxCapacity = v11; }).Ref(strAvailClassId, v12 => { strAvailClassId = v12; }).Ref(iStayNum, v13 => { iStayNum = v13; }).Ref(iThisReqmnt, v14 => { iThisReqmnt = v14; }).Ref(pO, v15 => { pO = v15; }));
                            }
                        }
                        if (_.IF(_.GT(_.NullableNUM(iLinkedUnitKey), (Int16)0)))
                        {
                            _.CALLm1v1(this, pO, "write", _.CONCAT("<input type=\"hidden\" name=\"linked_", iUnitKey, "\"  value=\"", iLinkedUnitKey, "\" />"));
                        }
                    }
                }
            }

            // If successfully received requirement data, complete form - otherwise render error
            if (_.IF(_.GT(_.NullableNUM(iThisReqmnt), (Int16)0)))
            {
                _.CALLm1v1(this, pO, "write", _.CONCAT("<input type=\"hidden\" name=\"availcal\" value=\"", _.CALLm0argp(this, _outer.Request, _.ARGS.Val("availcal")), "\" />"));
                _.CALLm1v1(this, pO, "write", _.CONCAT("<input type=\"hidden\" name=\"_nStays\" value=\"", iStayNum, "\" />"));
                _.CALLm1v1(this, pO, "write", _.CONCAT("<input type=\"hidden\" name=\"_nReqs\" value=\"", iThisReqmnt, "\" />"));

                // Close pnStayReqmntRslts div
                _.CALLm1v1(this, pO, "write", "</div>");

                _.CALLm1argp(this, _outer, "BookingUI_RenderButtons", _.ARGS.Ref(iStayNum, v16 => { iStayNum = v16; }).Ref(pO, v17 => { pO = v17; }).Val(false));

                // Close StayCandidateItem div
                _.CALLm1v1(this, pO, "write", "</div>");
            }
            else
            {
                _.CALLm1v1(this, pO, "Write", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/availcalendar/nounitsselectederror", "<h2>Error</h2><p class=\"error\">No units selected. Please click on the back button to return to the previous page and select the units you wish to book.</p>"));
            }

            // Close "staySelection" div and form
            _.CALLm1v1(this, pO, "Write", "</div>");
            _.CALLm1v1(this, pO, "Write", "</form>");
            if (_.IF(_.CALLm2v1(this, _outer.Page, "Site", "Params", "Booking_ChildPricing")))
            {
                _.CALLm1v1(this, pO, "Write", "<script type=\"text/javascript\">");
                _.CALLm1v1(this, pO, "Write", "NewMind.ETWP.Booking.UnitSelectionChildPricingGuests.Init();");
                _.CALLm1v1(this, pO, "Write", _.CONCAT("</", "script>"));
            }
            return BookingUI_StayMain_AvailCal_retVal;
        }
        // ====================================================================================================
        // RENDER: Main entry point when not using Availability Calendar
        // ====================================================================================================
        // SUMMARY: entry point for product UNIT/STAY selection (Booking)
        // [aiProductKey]: integer product key
        // [adtStartNight]: date of first night of stay
        // [aiNights]: integer number of nights
        // [aiFuzzyStayNumDays]: integer flexible start date days (ZERO = Precise match)
        public object BookingUI_StayMain(ref object objRenderSettings, ref object objData)
        {
            object BookingUI_StayMain_retVal = null;
            //most of these render functions rely on global variables, rather than trying to refactor them out for now ill create some globals
            //this needs refactoring

            // 2011-08-09 DWR: Expect the BookingRequirement in objRenderSettings to be read-only (since it usually comes from Page.Functions.GetSharedObject),
            // so replace it with an editable version (since some methods in here try to mess about with properties on it)
            _.SETm1a0(this, objRenderSettings, "BookingRequirement", _.OBJ(_.CALLm1v1(this, _outer, "GetEditableBookingRequirement", _.CALLm1v0(this, objRenderSettings, "BookingRequirement"))));

            _outer.IsVBPollingEnabled = _.VAL(_.CALLm1v0(this, objRenderSettings, "IsVBPollingEnabled"));
            _outer.bRenderAsCalendar = _.VAL(_.CALLm1v0(this, objRenderSettings, "RenderAsCalendar"));
            if (_.IF(_.IS(objData, VBScriptConstants.Nothing)))
            {
                // If couldn't retrieve product, report no availability - this will happen if the
                // availability criteria can (no longer) be met
                object byrefalias = objRenderSettings;
                try
                {
                    _.CALLm1argp(this, _outer, "RenderNoAvailElement", _.ARGS.Ref(byrefalias, v18 => { byrefalias = v18; }));
                }
                finally { objRenderSettings = byrefalias; }
                return BookingUI_StayMain_retVal;
            }

            if (_.IF(_.CALLm1v0(this, objRenderSettings, "LegacyRender")))
            {
                // Acco or Ticketing w/out VB Polling Enabled: Results from single Supplier (either
                // local OR FrontDesk for Acco, only local applies for Tickets)
                object byrefalias2 = objData, byrefalias3 = objRenderSettings;
                try
                {
                    _.CALLm1argp(this, _outer, "BookingUI_StayMain_Legacy", _.ARGS.Ref(byrefalias2, v19 => { byrefalias2 = v19; }).Ref(byrefalias3, v20 => { byrefalias3 = v20; }));
                }
                finally { objData = byrefalias2; objRenderSettings = byrefalias3; }
            }
            else
            {
                // Acco w/ VB Polling Enabled: Results from multiple Suppliers
                // - Not supported when handling Conference Bookings, these are local only (but when
                //   an OfferKey is set, IsVBPollingEnabled is put to False - see PreRender)
                object byrefalias4 = objData, byrefalias5 = objRenderSettings;
                try
                {
                    _.CALLm1argp(this, _outer, "BookingUI_StayMain_Polling", _.ARGS.Ref(byrefalias4, v21 => { byrefalias4 = v21; }).Ref(byrefalias5, v22 => { byrefalias5 = v22; }));
                }
                finally { objData = byrefalias4; objRenderSettings = byrefalias5; }
            }
            return BookingUI_StayMain_retVal;
        }
        // ====================================================================================================
        // RENDER: Write out form with hidden input fields used for internal or FrontDesk bookings
        // - This will open the form, but the caller must close it
        // ====================================================================================================
        // Note: We need to pass intProdKey into here as we may not have an objProduct reference
        // (eg. if called by BookingUI_StayMain_AvailCal)
        public object RenderBookingInfoForm(object pO, object intProdKey, object objRenderSettings, object intBookingType, object strSupplierId, object strSupplierName, object strSupplierEviivoName, object strSupplierDeepLinkQuality, object strSupplierLogo, object intEviivoSearchIndustryClassification)
        {
            object RenderBookingInfoForm_retVal = null;
            object strPostUrl = null;
            object strFormClass = null;
            object strNextStage = null;

            // intBookingType can be one of:
            //  BOOKING_Local => Proceed to "Checkout" stage next
            //  BOOKING_Eviivo => Proceed to "Checkout" stage next, but handling a FrontDesk
            //  BOOKING_Redirect => Will redirect to complete booking on separate (probably NewMind) site
            //  BOOKING_PollingRedirect => Will proceed to "PollingExit" stage next
            // strSupplierDeepLinkQuality should be null unless intBookingType is BOOKING_PollingRedirect,
            // in which case it should be a string (possibly null if we didn't have this information available
            // in the DMS about the Supplier)

            // Legacy rendering uses different form class for bookings that leave the site (we render them
            // all the same when VB Polling is enabled, though)
            if (_.IF(_.EQ(intBookingType, _outer.BOOKING_Redirect)))
            {
                strFormClass = "FrmUnitOptionsExt";
            }
            else
            {
                strFormClass = "FrmUnitOptions";
            }

            // What booking stage is next?
            // - If not external, go to checkout regardless of VB Polling setting.
            // - If IS external, branch off differently (VB Polling goes to a separate switcher stage, non-
            //   VB-Polling will redirect to the other site).
            // While we're here, retrieve POST url (secure for checkout, standard otherwise)
            if (_.IF(_.OR(_.EQ(intBookingType, _outer.BOOKING_Local), _.EQ(intBookingType, _outer.BOOKING_Eviivo))))
            {
                strNextStage = "checkout";
                strPostUrl = _.CONCAT(_.CALLm1v1(this, _outer, "GetPostUrl", true), "/", strNextStage);
            }
            else if (_.IF(_.EQ(intBookingType, _outer.BOOKING_Redirect)))
            {
                //strNextStage = "redirect"
                strNextStage = "checkout";
                //This should stay as "checkout" until 1.4 is updated to recognise "redirect" stage
                strPostUrl = _.VAL(_.CALLm2v1(this, _outer.Page, "PageInfo", "GetUrlFromPageID", "EXTBOOKPROMPT"));
                if (_.IF(_.ISNULL(strPostUrl)))
                {
                    _.CALLm1v1(this, _outer.Page, "PrintTraceWarning", "RenderBookingInfoForm: Unable to locate page EXTBOOKPROMPT, default to current page - is this correct behaviour??");
                    strPostUrl = _.VAL(_.CALLm2v0(this, _outer.Page, "URL", "Real"));
                }
            }
            else if (_.IF(_.EQ(intBookingType, _outer.BOOKING_PollingRedirect)))
            {
                // 2014-06-19 DWR: We have historically used the SupplierEviivoName for the URL segment, although it used to be labelled strSupplierName since
                // the values were getting set incorrectly. SupplierEviivoName seems like the most appropriate option since it will be a text-friendly string
                // value and so not have dots or spaces or whatever (and so be good for use in a URL).
                strNextStage = "pollingexit";
                strPostUrl = _.CONCAT(_.CALLm1v1(this, _outer, "GetPostUrl", false), "/pollingexit/", strSupplierEviivoName);
            }
            else
            {
                _.RAISEERROR(VBScriptConstants.vbObjectError, "ETWP.BookingUnitSelection", _.CONCAT("RenderBookingInfoForm: Invalid intBookingType value (", intBookingType, ")"));
            }

            _.CALLm1v1(this, pO, "Write", _.CONCAT("<form action=\"", strPostUrl, "\" "));
            if (_.IF(_.AND(_.NOT(_outer.IsVBPollingEnabled), _.EQ(_.NullableNUM(_.CALLm2v0(this, objRenderSettings, "BookingRequirement", "FlexibleRange")), (Int16)0))))
            {
                // Can't have ids when VB Polling enabled as we might be rendering out multiple of these forms.
                // 2008-11-10 DWR: This is similarly the case for fuzzy searching. I don't we have any working
                // Enterprise fuzzy-searching sites, so don't need to worry about breaking styling by removing
                // this id in this case.
                _.CALLm1v1(this, pO, "Write", "id=\"FrmUnitOptions\" ");
            }

            //#MJ's Reasoning -	In order for us to jump to unit selection in a tab it must have a name, however only the first form should have this
            if (_.IF(_.NOT(_outer.bFormRendered)))
            {
                _.CALLm1v1(this, pO, "Write", "name=\"FrmUnitOptions\" ");
                _outer.bFormRendered = true;
            }
            _.CALLm1v1(this, pO, "Write", _.CONCAT("class=\"", strFormClass, "\" method=\"post\">"));

            // Open container around common form values
            _.CALLm1v1(this, pO, "Write", "<div>");

            _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"stage\" value=\"", strNextStage, "\" />"));

            // Need to override market source if viewing site via widget
            if (_.IF(_.CALLm1v0(this, _outer.Page, "WidgetView")))
            {
                if (_.IF(_.EQ(intBookingType, _outer.BOOKING_Redirect)))
                {
                    // External bookings visit a preliminary redirect page first, which we want to be decluttered when in a widget
                    _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"widget_marketsource\" value=\"", _.CALLm1v0(this, _outer.Page, "WidgetMarketSource"), "\" />"));
                }
                else
                {
                    _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"msource\" value=\"", _.CALLm1v0(this, _outer.Page, "WidgetMarketSource"), "\" />"));
                }
                //this hidden field is to tell the checkout that weve come from a widget, and not a failed checkout validation
                _.CALLm1v1(this, pO, "Write", "<input type=\"hidden\" name=\"widget\" value=\"1\" />");
            }

            // None of this applies to VB Polling, even if it IS an external booking - we go to an
            // interim stage before leaving the site
            if (_.IF(_.EQ(intBookingType, _outer.BOOKING_Redirect)))
            {
                // NB: In "Conference Booking" mode (where OfferKey <> 0), we need to set the "channel" and "msource"
                //     values to different values (for msource, if there is no "ConfBookingMarketSourceID" set, it will
                //     fall back to using the site's main "MarketSourceID" source)
                _.CALLm1v1(this, pO, "Write", "<input type=\"hidden\" name=\"checkoutstage\" value=\"1\" />");
                if (_.IF(_.EQ(_.NullableNUM(_.CALLm2v0(this, objRenderSettings, "BookingRequirement", "Offer")), (Int16)0)))
                {
                    _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"channel\" value=\"", _.CALLm1v0(this, objRenderSettings, "Channel"), "\" />"));
                }
                else
                {
                    _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"channel\" value=\"", _.CALLm1v0(this, objRenderSettings, "ConfBookingChannel"), "\" />"));
                }
                if (_.IF(_.NOT(_.CALLm1v0(this, _outer.Page, "WidgetView"))))
                {
                    //Neeed to set market source override if redirecting to external site unless set above due to widgetview
                    if (_.IF(_.OR(_.EQ(_.NullableNUM(_.CALLm2v0(this, objRenderSettings, "BookingRequirement", "Offer")), (Int16)0), _.EQ(_.NullableSTR(_.CALLm2v1(this, _outer.Page, "Site", "Params", "ConfBookingMarketSourceID")), ""))))
                    {
                        _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"msource\" value=\"", _.CALLm2v1(this, _outer.Page, "Site", "Params", "MarketSourceID"), "\" />"));
                    }
                    else
                    {
                        _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"msource\" value=\"", _.CALLm2v1(this, _outer.Page, "Site", "Params", "ConfBookingMarketSourceID"), "\" />"));
                    }
                }
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"bookchannel\" value=\"", _.CALLm2v1(this, _outer.Page, "Site", "Params", "Booking_ChannelID"), "\" />"));
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"reposturl\" value=\"", _outer.strExtBookUrl, "\" />"));
                // 2009-09-21 DWR: New field to pass in so that the receiving site recognises booking as having
                // come from another site (so it can update appropriate Provider Stats)
                _.CALLm1v1(this, pO, "Write", "<input type=\"hidden\" name=\"ForcedExternalBooking\" value=\"1\" />");
            }

            _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"product\" value=\"", intProdKey, "\" />"));
            _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"isostartdate\" value=\"", _.CALLm3v1(this, _outer.Page, "Functions", "Dates", "ISODate", _.CALLm2v0(this, objRenderSettings, "BookingRequirement", "VisitDate")), "\" />"));
            _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"nights\" value=\"", _.CALLm2v0(this, objRenderSettings, "BookingRequirement", "Nights"), "\" />"));

            // We need all this when using VB Polling, even it it is an external booking, as we aren't
            // going to leave the site yet (there's an interim stage)
            if (_.IF(_.NOTEQ(intBookingType, _outer.BOOKING_Redirect)))
            {
                // NB: "package" parameter removed - it's now passed as "offer", and only when
                // customer is going for a "Conference Booking" discount product.
                // 2008-11-07 DWR: This used to referer to a "strRewriteUrl" value that was never defined.
                // So we'll pass in blank. Pretty sure it's not used anyway.
                _.CALLm1v1(this, pO, "Write", "<input type=\"hidden\" name=\"preUrl\" value=\"\" />");
                // 2008-11-07 DWR: If we've got non-precise results from a fuzzy search, we'll render this
                // form out and use the actual StartDate / NumNights combination that the fuzzy results
                // offered. So we just pass these to the checkout stage, and set "fuzzy" to zero.
                _.CALLm1v1(this, pO, "Write", "<input type=\"hidden\" name=\"fuzzy\" value=\"0\" />");
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"lng\" value=\"", _.CALLm2v0(this, _outer.Page, "Language", "LanguageCultureKey"), "\" />"));

                // NB: OfferKey is required for products in the "Conference Booking" functionality as
                // it lets the checkout object know that we should be looking for the product on the
                // "Conference Booking Channel" instead of the standard "website" channel. If this
                // ever needed to work with the ExternalBooking, we would need to pass out the
                // conference channel in the IsExternalBooking section above, but since this is
                // only being supported by the internal Newmind booking, it's not an issue.
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"offer\" value=\"", _.CALLm2v0(this, objRenderSettings, "BookingRequirement", "Offer"), "\" />"));

                // Pass in the current convert-to-currency value (this will have been held in the session
                // up to this point, but we may be about to leave the site when this form is posted, so
                // will need to send the value as a hidden input instead of relying on session)
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"CurrencyConvertTo\" value=\"", _.CALLm3v1(this, _outer.Page, "Functions", "Money", "GetCurrencyCodeOverride", _.CALLm2v0(this, _outer.Page, "Site", "LCCurrencyKey")), "\" />"));
            }

            // If we're dealing with a VB Polling External Supplier, write out the Supplier id, name and
            // deep-link-quality as well (this is the number of rooms that the supplier can handle in
            // deep-linking situations)
            if (_.IF(_.EQ(intBookingType, _outer.BOOKING_PollingRedirect)))
            {
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"SupplierId\" value=\"", strSupplierId, "\" />"));
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"SupplierName\" value=\"", strSupplierName, "\" />"));
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"SupplierLogo\" value=\"", strSupplierLogo, "\" />"));
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"SupplierEviivoName\" value=\"", strSupplierEviivoName, "\" />"));

                _.CALLm1v1(this, pO, "Write", "<input type=\"hidden\" name=\"EviivoSearchIndustryClassification\" value=\"");
                if (_.IF(_.ISNUMERIC(intEviivoSearchIndustryClassification)))
                {
                    _.CALLm1argp(this, pO, "Write", _.ARGS.Ref(intEviivoSearchIndustryClassification, v23 => { intEviivoSearchIndustryClassification = v23; }));
                }
                else
                {
                    _.CALLm1v1(this, pO, "Write", "0");
                }
                _.CALLm1v1(this, pO, "Write", "\" />");

                if (_.IF(_.ISNULL(strSupplierDeepLinkQuality)))
                {
                    strSupplierDeepLinkQuality = "";
                }
                else
                {
                    strSupplierDeepLinkQuality = _.VAL(_.TRIM(strSupplierDeepLinkQuality));
                }
                if (_.IF(_.NOT(_.ISNUMERIC(strSupplierDeepLinkQuality))))
                {
                    strSupplierDeepLinkQuality = "-1";
                }
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"SupplierDeepLinkQuality\" value=\"", strSupplierDeepLinkQuality, "\" />"));
            }

            // Append in the "Nominal Units" from Request collection or objUnitReqDictFromBookUrl (ie. "roomReq_1", "roomReq_2", etc..)
            //#MJ TODO need to call the new function
            _.CALLm1v1(this, pO, "Write", _.CALLm1v1(this, _outer, "GenerateRequirementFormData", _.CALLm1v0(this, objRenderSettings, "BookingRequirement")));
            // Close common form value container
            _.CALLm1v1(this, pO, "Write", "</div>");

            return RenderBookingInfoForm_retVal;
        }
        //generates a string of room requirement details in a format suitable for use in a form i.e hidden inputs ;)
        public object GenerateRequirementFormData(ref object objAccoSearchRequirement)
        {
            object GenerateRequirementFormData_retVal = null;
            object dictKeyValues = null;
            object aryFormattedData = null;
            object i = null;
            object key = null;
            //get our key value data dictionary
            object byrefalias6 = objAccoSearchRequirement;
            try
            {
                dictKeyValues = _.OBJ(_.CALLm3argp(this, _outer.Page, "Functions", "Booking", "GenerateRequirementKeyValueData", _.ARGS.Ref(byrefalias6, v24 => { byrefalias6 = v24; })));
            }
            finally { objAccoSearchRequirement = byrefalias6; }
            //create an array to hold our formatted data in which is the same size of the dictionary
            aryFormattedData = _.NEWARRAY(new object[] { _.SUBT(_.CALLm1v0(this, dictKeyValues, "Count"), (Int16)1) });
            //spin through our output array and add the formatted items in the format {key}={value}
            i = (Int16)0;
            var enumerationContent2 = _.ENUMERABLE(_.CALLm1v0(this, dictKeyValues, "Keys")).GetEnumerator();
            while (true)
            {
                if (!enumerationContent2.MoveNext())
                    break;
                key = enumerationContent2.Current;
                //#MJ's Reasoning -	we don't want to render our roomrequirement's here as they may not be valid
                //					instead when we write out a room requirement say what room it is and for how many
                // NP: MJ is saying add hidden form values based on the requirements linked to the UnitStayDetails by the AvailCom
                // NOT to base it on the BookingRequestDictionary. These form values will then be posted
                // and update the BookingRequirement object for when it is used in the Booking Checkout
                if (_.IF(_.NOT(_.EQ(_.NullableSTR(_.LEFT(_.LCASE(key), (Int16)8)), "roomreq_"))))
                {
                    _.SETm0a1(this, aryFormattedData, i, _.CONCAT("<input type=\"hidden\" name=\"", key, "\" value=\"", _.CALLm1argp(this, dictKeyValues, "Item", _.ARGS.Ref(key, v26 => { key = v26; })), "\" />", VBScriptConstants.vbCrLf));
                }
                i = _.ADD(i, (Int16)1);
            }
            //return our array as a string using an & as the joining character
            GenerateRequirementFormData_retVal = _.JOIN(aryFormattedData);
            return GenerateRequirementFormData_retVal;
        }
        public object GetPostUrl(object bSecure)
        {
            object GetPostUrl_retVal = null;
            object strPostUrl = null;
            object strUrl = null;

            if (_.IF(bSecure))
            {
                strPostUrl = _.VAL(_.CALLm2v0(this, _outer.Page, "Site", "SecureHostName"));
            }
            else
            {
                strPostUrl = _.VAL(_.CALLm2v0(this, _outer.Page, "URL", "FullHostName"));
            }
            while (_.IF(_.EQ(_.NullableSTR(_.RIGHT(strPostUrl, (Int16)1)), "/")))
            {
                strPostUrl = _.VAL(_.LEFT(strPostUrl, _.SUBT(_.LEN(strPostUrl), (Int16)1)));
            }

            strUrl = _.VAL(_.CALLm2v1(this, _outer.Page, "PageInfo", "GetUrlFromPageID", "BOOKONLINE"));
            if (_.IF(_.ISNULL(strUrl)))
            {
                _.CALLm1v1(this, _outer.Page, "PrintTraceWarning", "GetPostUrl: Unable to locate page BOOKONLINE, default to current page - is this correct behaviour??");
                strUrl = _.VAL(_.CALLm2v0(this, _outer.Page, "URL", "Real"));
            }

            if (_.IF(_.NOTEQ(_.NullableSTR(_.LEFT(strUrl, (Int16)1)), "/")))
            {
                strUrl = _.CONCAT("/", strUrl);
            }

            strPostUrl = _.CONCAT(strPostUrl, strUrl);
            if (_.IF(_.AND(_.NOTEQ(_.NullableSTR(_.UCASE(_.LEFT(strPostUrl, (Int16)7))), "HTTP://"), _.NOTEQ(_.NullableSTR(_.UCASE(_.LEFT(strPostUrl, (Int16)8))), "HTTPS://"))))
            {
                strPostUrl = _.CONCAT("http://", strPostUrl);
            }

            GetPostUrl_retVal = _.VAL(strPostUrl);
            return GetPostUrl_retVal;
        }
        // SUMMARY: render new requirement UI from avail calendar
        // [ireqSz]: ADO unit recordset from availability object
        // [aiStayNum]: integer stay index
        // [aiThisReqmnt]: integer requirement number (from recordset)
        public object BookingUI_RenderNewReq_AvailCal(ref object intBookingType, ref object iUnitKey, ref object strUnitName, ref object iUnitMinOccupancy, ref object iUnitMaxCapacity, ref object asAvailClassId, ref object aiStayNum, ref object aiThisReqmnt, ref object pO)
        {
            object BookingUI_RenderNewReq_AvailCal_retVal = null;
            object iGuest = null;
            object strGuestsFor = null;
            object strAdultsTitle = null;
            object strAdults = null;
            object strChildrenTitle = null;
            object strChildren = null;
            object strGuestsAnd = null;
            object iCount = null;
            object ageValue = null;
            object strGuestsTitle = null;
            object strGuests = null;
            // on first ever call [aiThisReqmnt]=1, on subsequent calls we must close previous [pnStayReqmnt] and [pnStayReqmntRslts] DIVs

            if (_.IF(_.GT(_.NullableNUM(aiThisReqmnt), (Int16)1)))
            {
                _.CALLm1v1(this, pO, "Write", "</div></div>");
            }

            _.CALLm1v1(this, pO, "Write", _.CONCAT("<div class=\"pnStayReqmnt\">", VBScriptConstants.vbCrLf));
            _.CALLm1v1(this, pO, "Write", _.CONCAT("<div class=\"pnStayReqmntTtl\">", VBScriptConstants.vbCrLf));
            object byrefalias7 = asAvailClassId;
            try
            {
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<div Class=\"pnStayReqmntRoom\">Room ", aiThisReqmnt, " - ", strUnitName, _.CALLm1argp(this, _outer, "BookingUI_AvailClassIcon", _.ARGS.Ref(byrefalias7, v28 => { byrefalias7 = v28; })), " <br/></div>"));
            }
            finally { asAvailClassId = byrefalias7; }

            if (_.IF(_.OR(_.EQ(_.NullableNUM(iUnitMinOccupancy), (Int16)0), _.EQ(_.NullableSTR(iUnitMinOccupancy), ""))))
            {
                iUnitMinOccupancy = (Int16)1;
            }

            _.CALLm1v1(this, pO, "Write", _.CONCAT("<div Class=\"pnStayReqmntGuests\">", VBScriptConstants.vbCrLf));
            if (_.IF(_.EQ(iUnitMaxCapacity, iUnitMinOccupancy)))
            {
                _.CALLm1v1(this, pO, "Write", _.CONCAT("For ", iUnitMaxCapacity, " guests <input type=\"hidden\" name=\"roomReq_", aiThisReqmnt, "\" value=\"", iUnitMaxCapacity, "\"/>"));
            }
            else
            {
                strGuestsFor = _.VAL(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/guestrequirement/for", "for"));
                //alas child pricing is different
                if (_.IF(_.CALLm2v1(this, _outer.Page, "Site", "Params", "Booking_ChildPricing")))
                {

                    strAdultsTitle = _.VAL(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/guestrequirement/adults/selecttitle", "Please specify the number of adults in this room."));
                    strAdults = _.VAL(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/guestrequirement/adults/adult(s)", "adult(s)"));

                    _.CALLm1v1(this, pO, "Write", _.CONCAT(strGuestsFor, " <select class=\"adults\" name=\"roomReq_", aiThisReqmnt, "_adults\" title=\"", strAdultsTitle, "\"> "));
                    var loopEnd2 = _.NUM(iUnitMaxCapacity);
                    var loopStart2 = _.NUM(iUnitMinOccupancy, loopEnd2, (Int16)1);
                    if (_.StrictLTE(loopStart2, loopEnd2))
                    {
                        for (iGuest = loopStart2; _.StrictLTE(iGuest, loopEnd2); iGuest = _.ADD(iGuest, (Int16)1))
                        {
                            _.CALLm1v1(this, pO, "Write", _.CONCAT("<option value=\"", iGuest, "\">", iGuest, "</option> "));
                        }
                    }
                    _.CALLm1v1(this, pO, "Write", _.CONCAT("</select> ", strAdults));

                    strChildrenTitle = _.VAL(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/guestrequirement/children/selecttitle", "Please specify the number of children in this room."));
                    strChildren = _.VAL(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/guestrequirement/children/children", "children"));
                    strGuestsAnd = _.VAL(_.CALLm1v2(this, _outer.Page, "Resource", "and", "and"));
                    _.CALLm1v1(this, pO, "Write", _.CONCAT(" ", strGuestsAnd, " <select class=\"children\" name=\"roomReq_", aiThisReqmnt, "_children\" title=\"", strChildrenTitle, "\"> "));

                    var loopEnd3 = _.NUM(_.SUBT(iUnitMaxCapacity, (Int16)1));
                    var loopStart3 = _.NUM((Int16)0, loopEnd3, (Int16)1);
                    if (_.StrictLTE(loopStart3, loopEnd3))
                    {
                        for (iGuest = loopStart3; _.StrictLTE(iGuest, loopEnd3); iGuest = _.ADD(iGuest, (Int16)1))
                        {
                            _.CALLm1v1(this, pO, "Write", _.CONCAT("<option value=\"", iGuest, "\">", iGuest, "</option> "));
                        }
                    }
                    _.CALLm1v1(this, pO, "Write", _.CONCAT("</select> ", strChildren));

                    _.CALLm1v1(this, pO, "WriteLine", "<span class=\"label childrenageslabel\">Child Ages</span>");
                    _.CALLm1v1(this, pO, "WriteLine", "<span class=\"field childrenagesfield\">");

                    var loopEnd4 = _.NUM(_.SUBT(iUnitMaxCapacity, (Int16)1));
                    var loopStart4 = _.NUM((Int16)0, loopEnd4, (Int16)1);
                    if (_.StrictLTE(loopStart4, loopEnd4))
                    {
                        for (iCount = loopStart4; _.StrictLTE(iCount, loopEnd4); iCount = _.ADD(iCount, (Int16)1))
                        {
                            _.CALLm1v1(this, pO, "WriteLine", "<span class=\"childageWrapper\">");
                            _.CALLm1v1(this, pO, "WriteLine", _.CONCAT(VBScriptConstants.vbTab, "<span class=\"label childagelabel\">Child Age ", _.ADD(iCount, (Int16)1), "</span>"));
                            _.CALLm1v1(this, pO, "WriteLine", _.CONCAT(VBScriptConstants.vbTab, "<span class=\"field childagefield\">"));
                            _.CALLm1v1(this, pO, "Write", _.CONCAT("<select class=\"\" name=\"roomReq_", aiThisReqmnt, "_children_childage", iCount, "\">"));
                            for (iGuest = (Int16)0; _.StrictLTE(iGuest, 18); iGuest = _.ADD(iGuest, (Int16)1))
                            {
                                _.CALLm1v1(this, pO, "Write", _.CONCAT("<option value=\"", iGuest, "\">", iGuest, "</option> "));
                            }
                            _.CALLm1v1(this, pO, "Write", "</select> ");
                            _.CALLm1v1(this, pO, "WriteLine", _.CONCAT(VBScriptConstants.vbTab, "</span>"));
                            _.CALLm1v1(this, pO, "WriteLine", "</span>");
                        }
                    }
                    _.CALLm1v1(this, pO, "WriteLine", "</span>");
                }
                else
                {
                    strGuestsTitle = _.VAL(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/guestrequirement/selecttitle", "Please specify the number of guests in this room."));
                    strGuests = _.VAL(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/guestrequirement/guest(s)", "guest(s)"));

                    _.CALLm1v1(this, pO, "Write", _.CONCAT(strGuestsFor, " <select name=\"roomReq_", aiThisReqmnt, "\" title=\"", strGuestsTitle, "\"> "));
                    var loopEnd5 = _.NUM(iUnitMaxCapacity);
                    var loopStart5 = _.NUM(iUnitMinOccupancy, loopEnd5, (Int16)1);
                    if (_.StrictLTE(loopStart5, loopEnd5))
                    {
                        for (iGuest = loopStart5; _.StrictLTE(iGuest, loopEnd5); iGuest = _.ADD(iGuest, (Int16)1))
                        {
                            _.CALLm1v1(this, pO, "Write", _.CONCAT("<option value=\"", iGuest, "\">", iGuest, "</option> "));
                        }
                    }
                    _.CALLm1v1(this, pO, "Write", _.CONCAT("</select> ", strGuests));
                }
            }
            _.CALLm1v1(this, pO, "Write", _.CONCAT("</div>", VBScriptConstants.vbCrLf));

            if (_.IF(_.EQ(_.NullableSTR(intBookingType), "ticketing")))
            {
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"unit_", iUnitKey, "\"  value=\"", aiThisReqmnt, "\" />"));
            }
            else
            {
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"unit_", aiStayNum, "_", aiThisReqmnt, "\"  value=\"", iUnitKey, "\" />"));
            }
            return BookingUI_RenderNewReq_AvailCal_retVal;
        }
        // SUMMARY: Draw availability month calendar
        // [sbCalendars]:  ASP [nmStringBuilder] object instance output string
        // [dCalStartDflt]: date default calendar start date
        // <retval>: string month available stays details JSON data
        public object BookingUI_RenderAvailCal(ref object sbCalendars, ref object objDictAvaiStays, ref object bStarted)
        {
            object BookingUI_RenderAvailCal_retVal = null;
            object strClassMonth = null;
            object dStart1 = null;
            object aryAvailStaysKeys = null;

            strClassMonth = "MonthWrapper";
            if (_.IF(_.NOT(bStarted)))
            {
                _.CALLm1v1(this, sbCalendars, "AppendLine", "<div class=\"CalendarsWrapper\">");
                _.CALLm1v1(this, sbCalendars, "AppendLine", _.CONCAT("<div class=\"instruction\">", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/availcalendar/instruction", "Please select an available stay from the calendars below. Clicking on a highlighted start day for a stay will show the stay details such as the units available, price, etc."), "</div>"));
                strClassMonth = _.CONCAT(strClassMonth, " currentmonth");
            }
            else
            {
                strClassMonth = _.CONCAT(strClassMonth, " nextmonth");
            }

            if (_.IF(_.GT(_.NullableNUM(_.CALLm1v0(this, objDictAvaiStays, "Count")), (Int16)0)))
            {
                aryAvailStaysKeys = _.VAL(_.CALLm1v0(this, objDictAvaiStays, "Keys"));
                dStart1 = _.REPLACE(_.CALLm0argp(this, aryAvailStaysKeys, _.ARGS.Val((Int16)0)), "sd_", "");
                _.ERASE(aryAvailStaysKeys, v30 => { aryAvailStaysKeys = v30; });
            }
            else
            {
                dStart1 = _.DATE();
            }

            object byrefalias8 = sbCalendars, byrefalias9 = objDictAvaiStays;
            try
            {
                _.CALLm1argp(this, _outer, "BookingUI_RenderCalendarMonthWithAvailability", _.ARGS.Ref(byrefalias8, v31 => { byrefalias8 = v31; }).Ref(dStart1, v32 => { dStart1 = v32; }).Ref(strClassMonth, v33 => { strClassMonth = v33; }).Ref(byrefalias9, v34 => { byrefalias9 = v34; }));
            }
            finally { sbCalendars = byrefalias8; objDictAvaiStays = byrefalias9; }

            // using a global count so we can track how many calendars have been added to the stringbuilder for the prev/next buttons
            // doing this now because of the recursive nature of this function
            _outer.g_iNumberOfCalendarsRendered = _.ADD(_outer.g_iNumberOfCalendarsRendered, (Int16)1);

            //Check if we have stays left and render then as another calendar
            if (_.IF(_.GT(_.NullableNUM(_.CALLm1v0(this, objDictAvaiStays, "Count")), (Int16)0)))
            {
                object byrefalias10 = sbCalendars, byrefalias11 = objDictAvaiStays;
                try
                {
                    _.CALLm1argp(this, _outer, "BookingUI_RenderAvailCal", _.ARGS.Ref(byrefalias10, v35 => { byrefalias10 = v35; }).Ref(byrefalias11, v36 => { byrefalias11 = v36; }).Val(true));
                }
                finally { sbCalendars = byrefalias10; objDictAvaiStays = byrefalias11; }
            }
            else
            {
                //not sure if this should be dStart1 - was dStart
                object byrefalias12 = sbCalendars;
                try
                {
                    _.CALLm1argp(this, _outer, "BookingUI_RenderAvailCalLinks", _.ARGS.Ref(dStart1, v37 => { dStart1 = v37; }).Ref(byrefalias12, v38 => { byrefalias12 = v38; }));
                }
                finally { sbCalendars = byrefalias12; }
                object byrefalias13 = sbCalendars;
                try
                {
                    _.CALLm1argp(this, _outer, "BookingUI_RenderAvailCalKey", _.ARGS.Ref(byrefalias13, v39 => { byrefalias13 = v39; }));
                }
                finally { sbCalendars = byrefalias13; }
                _.CALLm1v1(this, sbCalendars, "AppendLine", "</div>");

            }

            return BookingUI_RenderAvailCal_retVal;
        }
        public object BookingUI_RenderCalendarMonth(ref object sbCalendars, object dFirstDayOfMonth, object strWrapperClass)
        {
            object BookingUI_RenderCalendarMonth_retVal = null;
            object byrefalias14 = sbCalendars;
            try
            {
                _.CALLm1argp(this, _outer, "BookingUI_RenderCalendarMonthWithAvailability", _.ARGS.Ref(byrefalias14, v40 => { byrefalias14 = v40; }).Ref(dFirstDayOfMonth, v41 => { dFirstDayOfMonth = v41; }).Ref(strWrapperClass, v42 => { strWrapperClass = v42; }).Val(VBScriptConstants.Nothing));
            }
            finally { sbCalendars = byrefalias14; }
            return BookingUI_RenderCalendarMonth_retVal;
        }
        public object BookingUI_RenderCalendarMonthWithAvailability(ref object sbCalendars, object dFirstDayOfMonth, object strWrapperClass, object objDictAvailStays)
        {
            object BookingUI_RenderCalendarMonthWithAvailability_retVal = null;
            object iWeekStartDay = null;
            object iWeekDayCalStart = null;
            object iWeekDayCalEnd = null;
            object dCalStart = null;
            object dCalEnd = null;
            object strThisMonthYear = null;
            object strTableSummary = null;
            object strHeaderCellClass = null;
            object i = null;
            object iCellCount = null;
            object bFirstCell = null;
            object bLastCell = null;
            object dDate = null;
            object bStartNewStay = null;
            object bStayIndicative = null;
            object strStayNumber = null;
            object iDay = null;
            object iPrePadding = null;
            object j = null;
            object strDisplayText = null;
            object strDayCellClass = null;
            object aryStay = null;
            object strAvailType = null;
            object strIndicativeIcon = null;
            object iPostPadding = null;
            object k = null;

            iWeekStartDay = (Int16)1; //Monday
            iWeekDayCalStart = _.MOD(_.ADD(iWeekStartDay, (Int16)1), (Int16)7);
            iWeekDayCalEnd = _.MOD(iWeekStartDay, (Int16)7);

            dCalStart = _.VAL(_.CALLm3argp(this, _outer.Page, "Functions", "Dates", "fn_GetFirstDateOfMonth", _.ARGS.Ref(dFirstDayOfMonth, v43 => { dFirstDayOfMonth = v43; })));
            dCalEnd = _.VAL(_.CALLm3argp(this, _outer.Page, "Functions", "Dates", "fn_GetLastDateOfMonth", _.ARGS.Ref(dFirstDayOfMonth, v44 => { dFirstDayOfMonth = v44; })));
            strThisMonthYear = _.CONCAT(_.CALLm3v1(this, _outer.Page, "Functions", "Dates", "GetMonthNameAbbr", _.MONTH(dCalStart)), " ", _.YEAR(dCalStart));
            strTableSummary = _.CONCAT(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/availcalendar/availabilitycalendarfor", "Availability calendar for"), " ", strThisMonthYear);

            _.CALLm1v1(this, sbCalendars, "AppendLine", _.CONCAT("<div id=\"Cal_", _.CALLm3argp(this, _outer.Page, "Functions", "Dates", "ISODate", _.ARGS.Ref(dCalStart, v45 => { dCalStart = v45; })), "\" class=\"", strWrapperClass, "\">"));
            _.CALLm1v1(this, sbCalendars, "AppendLine", _.CONCAT("<table id=\"Tbl_", _.CALLm3argp(this, _outer.Page, "Functions", "Dates", "ISODate", _.ARGS.Ref(dCalStart, v47 => { dCalStart = v47; })), "\" class=\"availabilityCalendar\" summary=\"", strTableSummary, "\" >"));
            _.CALLm1v1(this, sbCalendars, "AppendLine", "<thead>");
            _.CALLm1v1(this, sbCalendars, "AppendLine", "<tr>");
            _.CALLm1v1(this, sbCalendars, "AppendLine", _.CONCAT("<th colspan=\"8\">", strThisMonthYear, "</th>"));
            _.CALLm1v1(this, sbCalendars, "AppendLine", "</tr>");

            strHeaderCellClass = "";
            var loopEnd6 = _.NUM(_.ADD(iWeekStartDay, (Int16)6));
            var loopStart6 = _.NUM(iWeekStartDay, loopEnd6, (Int16)1);
            if (_.StrictLTE(loopStart6, loopEnd6))
            {
                for (i = loopStart6; _.StrictLTE(i, loopEnd6); i = _.ADD(i, (Int16)1))
                {
                    if (_.IF(_.OR(_.EQ(_.NullableNUM(_.MOD(i, (Int16)7)), (Int16)6), _.EQ(_.NullableNUM(_.MOD(i, (Int16)7)), (Int16)0))))
                    {
                        strHeaderCellClass = " class=\"we\"";
                    }
                    _.CALLm1v1(this, sbCalendars, "AppendLine", _.CONCAT("<th", strHeaderCellClass, ">", _.CALLm3v1(this, _outer.Page, "Functions", "Dates", "GetDayNameAbbr", _.WEEKDAY(_.MOD(_.ADD(i, (Int16)1), (Int16)7))), "</th>"));
                }
            }

            _.CALLm1v1(this, sbCalendars, "AppendLine", "</tr>");
            _.CALLm1v1(this, sbCalendars, "AppendLine", "</thead>");
            _.CALLm1v1(this, sbCalendars, "AppendLine", "<tbody>");

            iCellCount = (Int16)0;
            bFirstCell = true;
            bLastCell = false;

            dDate = _.VAL(dCalStart);

            var loopEnd7 = _.NUM(_.DAY(dCalEnd));
            var loopStart7 = _.NUM(_.DAY(dCalStart), loopEnd7, (Int16)1);
            if (_.StrictLTE(loopStart7, loopEnd7))
            {
                for (iDay = loopStart7; _.StrictLTE(iDay, loopEnd7); iDay = _.ADD(iDay, (Int16)1))
                {
                    bStartNewStay = false;

                    if (_.IF(bFirstCell))
                    {
                        iPrePadding = _.VAL(_.DATEDIFF("d", _.CALLm3argp(this, _outer.Page, "Functions", "Dates", "fn_GetFirstDateOfWeek", _.ARGS.Val(_.CALLm3argp(this, _outer.Page, "Functions", "Dates", "fn_GetFirstDateOfMonth", _.ARGS.Ref(dCalStart, v49 => { dCalStart = v49; }))).Ref(iWeekDayCalStart, v50 => { iWeekDayCalStart = v50; })), dCalStart));
                        if (_.IF(_.GT(_.NullableNUM(iPrePadding), (Int16)0)))
                        {
                            var loopEnd8 = _.NUM(iPrePadding);
                            var loopStart8 = _.NUM((Int16)1, loopEnd8);
                            if (_.StrictLTE(loopStart8, loopEnd8))
                            {
                                for (j = loopStart8; _.StrictLTE(j, loopEnd8); j = _.ADD(j, (Int16)1))
                                {
                                    _.CALLm1v1(this, sbCalendars, "AppendLine", "<td></td>");
                                    iCellCount = _.ADD(iCellCount, (Int16)1);
                                }
                            }
                        }
                        bFirstCell = false;
                    }

                    strDisplayText = _.CONCAT("", iDay);
                    strDayCellClass = "n";

                    if (_.IF(_.NOT(_.IS(objDictAvailStays, VBScriptConstants.Nothing))))
                    {
                        if (_.IF(_.CALLm1v1(this, objDictAvailStays, "Exists", _.CONCAT("sd_", dDate))))
                        {
                            bStartNewStay = true;
                            //we expect value in the format [stayNo]_[indicative]
                            aryStay = _.SPLIT(_.CALLm0argp(this, objDictAvailStays, _.ARGS.Val(_.CONCAT("sd_", dDate))), "_");
                            strStayNumber = _.VAL(_.CALLm0argp(this, aryStay, _.ARGS.Val((Int16)0)));
                            bStayIndicative = _.CBOOL(_.CALLm0argp(this, aryStay, _.ARGS.Val((Int16)1)));
                            _.CALLm1v1(this, objDictAvailStays, "Remove", _.CONCAT("sd_", dDate));
                            _.ERASE(aryStay, v51 => { aryStay = v51; });
                        }
                    }

                    if (_.IF(_.LT(dDate, _.DATE()))) //date is in the past
                    {

                        strDayCellClass = "p";

                    }
                    else if (_.IF(bStartNewStay))
                    {

                        if (_.IF(_.NOT(_.IS(objDictAvailStays, VBScriptConstants.Nothing))))
                        {
                            strAvailType = "";
                            strIndicativeIcon = "";

                            if (_.IF(bStayIndicative))
                            {
                                strDayCellClass = "i";
                                strAvailType = _.VAL(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/unconfirmedavailability", "Unconfirmed Availability"));
                                strIndicativeIcon = _.CONCAT("<img src=\"", _.CALLm1v2(this, _outer.Page, "ImageResource", "bookonline/icons/indicative", "/images/icon_indicative.gif"), "\" alt=\"", strAvailType, "\" class=\"icon\"/>");
                            }
                            else
                            {
                                strDayCellClass = "a";
                                strAvailType = _.VAL(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/confirmedavailability", "Confirmed Availability"));
                                strIndicativeIcon = _.CONCAT("<img src=\"", _.CALLm1v2(this, _outer.Page, "ImageResource", "bookonline/icons/allocated", "/images/icon_allocated.gif"), "\" alt=\"", strAvailType, "\" class=\"icon\"/>");
                            }

                            strDisplayText = _.CONCAT("<a href=\"#stay_", strStayNumber, "\" class=\"calavailstay\" id=\"stay_", strStayNumber, "\">", _.DAY(dDate), "</a>", strIndicativeIcon);

                        }

                    }

                    if (_.IF(_.OR(_.EQ(_.NullableNUM(_.WEEKDAY(dDate)), (Int16)1), _.EQ(_.NullableNUM(_.WEEKDAY(dDate)), (Int16)7))))
                    {
                        strDayCellClass = _.CONCAT(strDayCellClass, " we");
                    }

                    _.CALLm1v1(this, sbCalendars, "AppendLine", _.CONCAT("<td class=\"", strDayCellClass, "\"><div>", strDisplayText, "</div></td>"));

                    iCellCount = _.ADD(iCellCount, (Int16)1);

                    if (_.IF(_.EQ(dDate, dCalEnd)))
                    {
                        bLastCell = true;
                    }

                    // This is for when the last day of the month is not the last day of the week and empty cells are put in place to fill the calendar days
                    if (_.IF(bLastCell))
                    {
                        iPostPadding = _.VAL(_.DATEDIFF("d", dCalEnd, _.CALLm3argp(this, _outer.Page, "Functions", "Dates", "fn_GetLastDateOfWeek", _.ARGS.Ref(dCalEnd, v52 => { dCalEnd = v52; }).Ref(iWeekDayCalEnd, v53 => { iWeekDayCalEnd = v53; }))));
                        if (_.IF(_.AND(_.GT(_.NullableNUM(iPostPadding), (Int16)0), _.LT(_.NullableNUM(iPostPadding), (Int16)7))))
                        {
                            var loopEnd9 = _.NUM(iPostPadding);
                            var loopStart9 = _.NUM((Int16)1, loopEnd9);
                            if (_.StrictLTE(loopStart9, loopEnd9))
                            {
                                for (k = loopStart9; _.StrictLTE(k, loopEnd9); k = _.ADD(k, (Int16)1))
                                {
                                    _.CALLm1v1(this, sbCalendars, "AppendLine", "<td></td>");
                                    iCellCount = _.ADD(iCellCount, (Int16)1);
                                }
                            }
                        }
                        bLastCell = false;
                        bFirstCell = true;
                    }

                    if (_.IF(_.EQ(_.NullableNUM(_.MOD(iCellCount, (Int16)7)), (Int16)0)))
                    {
                        _.CALLm1v1(this, sbCalendars, "AppendLine", "</tr>");
                    }

                    dDate = _.VAL(_.DATEADD("d", 1, dDate));
                }
            }

            _.CALLm1v1(this, sbCalendars, "AppendLine", "</tbody>");
            _.CALLm1v1(this, sbCalendars, "AppendLine", "</table>");
            _.CALLm1v1(this, sbCalendars, "AppendLine", "</div>");

            return BookingUI_RenderCalendarMonthWithAvailability_retVal;
        }
        public object BookingUI_RenderAvailCalKey(ref object sb)
        {
            object BookingUI_RenderAvailCalKey_retVal = null;
            object strCalKey = null;
            strCalKey = _.VAL(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/availcalendar/calkey", ""));
            if (_.IF(_.NOTEQ(_.NullableSTR(_.TRIM(strCalKey)), "")))
            {
                _.CALLm1v1(this, sb, "AppendLine", _.CONCAT("<div class=\"CalKey\">", strCalKey, "</div>"));
            }
            return BookingUI_RenderAvailCalKey_retVal;
        }
        public object BookingUI_RenderAvailCalLinks(ref object dStart, ref object sb)
        {
            object BookingUI_RenderAvailCalLinks_retVal = null;
            object dCalStartPrev = null;
            object strTitlePrev = null;
            object dCalStartNext = null;
            object strTitleNext = null;
            object iPositiveMonthAdjustment = null;
            object iNegativeMonthAdjustment = null;

            // dStart is the start date for the last month shown in the rendered calendars
            // and we therefore only need to go forward by 1 month
            // even if no calendars are shown for the current month we can still potentially
            // move to a future month where there is availability.
            iPositiveMonthAdjustment = (Int16)1;
            // The previous month link has to go back by however many months are already showing, i.e. Jul & Aug are shown
            // dStart = 01/08/2011 (Aug) and we need to display Jun & Jul so we need to jump back 2 months to June.
            iNegativeMonthAdjustment = _.SUBT(_outer.g_iNumberOfCalendarsRendered);

            if (_.IF(_.EQ(_.NullableNUM(_outer.g_iNumberOfCalendarsRendered), (Int16)0)))
            {
                // If we have no rendered calendars we still need the link to go back by 1 month
                iNegativeMonthAdjustment = (Int16)(-1);
            }

            dCalStartPrev = _.VAL(_.CALLm3v1(this, _outer.Page, "Functions", "Dates", "fn_GetFirstDateOfMonth", _.DATEADD("m", iNegativeMonthAdjustment, dStart)));
            strTitlePrev = _.VAL(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/availcalendar/previousmonth", "&lt;&lt; Previous Month"));

            dCalStartNext = _.VAL(_.CALLm3v1(this, _outer.Page, "Functions", "Dates", "fn_GetFirstDateOfMonth", _.DATEADD("m", iPositiveMonthAdjustment, dStart)));
            strTitleNext = _.VAL(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/availcalendar/nextmonth", "Next Month &gt;&gt;"));

            _.CALLm1v1(this, sb, "AppendLine", "<div class=\"CalNavLinks\">");
            _.CALLm1v1(this, sb, "AppendLine", _.CALLm1argp(this, _outer, "BookingUI_RenderAvailCalLink", _.ARGS.Ref(dCalStartPrev, v54 => { dCalStartPrev = v54; }).Ref(strTitlePrev, v55 => { strTitlePrev = v55; }).Val("prev")));
            _.CALLm1v1(this, sb, "AppendLine", _.CALLm1argp(this, _outer, "BookingUI_RenderAvailCalLink", _.ARGS.Ref(dCalStartNext, v56 => { dCalStartNext = v56; }).Ref(strTitleNext, v57 => { strTitleNext = v57; }).Val("next")));
            _.CALLm1v1(this, sb, "AppendLine", "</div>");

            return BookingUI_RenderAvailCalLinks_retVal;
        }
        public object BookingUI_RenderAvailCalLink(ref object dCalStartDate, ref object strTitle, ref object strClass)
        {
            object BookingUI_RenderAvailCalLink_retVal = null;
            object itm = null;
            object sValue = null;
            object strLink = null;
            object bFound = null;

            bFound = false;

            var enumerationContent3 = _.ENUMERABLE(_.CALLm1v0(this, _outer.Request, "QueryString")).GetEnumerator();
            while (true)
            {
                if (!enumerationContent3.MoveNext())
                    break;
                itm = enumerationContent3.Current;
                if (_.IF(_.EQ(_.NullableSTR(itm), "isostartdate")))
                {
                    //reset date
                    object byrefalias15 = dCalStartDate;
                    try
                    {
                        sValue = _.VAL(_.CALLm3argp(this, _outer.Page, "Functions", "Dates", "ISODate", _.ARGS.Ref(byrefalias15, v58 => { byrefalias15 = v58; })));
                    }
                    finally { dCalStartDate = byrefalias15; }
                    bFound = true;
                }
                else
                {
                    sValue = _.VAL(_.CALLm1argp(this, _outer.Request, "QueryString", _.ARGS.Ref(itm, v59 => { itm = v59; })));
                }
                strLink = _.CONCAT(strLink, "&amp;", itm, "=", _.CALLm1argp(this, _outer.Server, "UrlEncode", _.ARGS.Ref(sValue, v60 => { sValue = v60; })));
            }

            if (_.IF(_.NOT(bFound)))
            {
                object byrefalias16 = dCalStartDate;
                try
                {
                    strLink = _.CONCAT(strLink, "&amp;isostartdate=", _.CALLm1v1(this, _outer.Server, "UrlEncode", _.CALLm3argp(this, _outer.Page, "Functions", "Dates", "ISODate", _.ARGS.Ref(byrefalias16, v62 => { byrefalias16 = v62; }))));
                }
                finally { dCalStartDate = byrefalias16; }
            }

            if (_.IF(_.NOTEQ(_.NullableSTR(_.TRIM(_.CONCAT("", strLink))), "")))
            {
                strLink = _.REPLACE(strLink, "&amp;", "?", (Int16)1, (Int16)1, (Int16)0);
            }

            if (_.IF(_.GTE(_.NullableNUM(_.DATEDIFF("m", _.DATE(), dCalStartDate)), (Int16)0)))
            {
                BookingUI_RenderAvailCalLink_retVal = _.CONCAT("<a href=\"", strLink, "\" class=\"", strClass, "\" title=\"", strTitle, "\" rel=\"nofollow\">", strTitle, "</a>", VBScriptConstants.vbCrLf);
            }
            else
            {
                BookingUI_RenderAvailCalLink_retVal = "";
            }

            return BookingUI_RenderAvailCalLink_retVal;
        }
        // ====================================================================================================
        // RENDER: Main entry point when VB Polling is enabled
        // - Applies to acco products only
        // - Not supported when handling Conference Bookings (these are local only)
        // ====================================================================================================
        public object BookingUI_StayMain_Polling(ref object objData, ref object objRenderSettings)
        {
            object BookingUI_StayMain_Polling_retVal = null;
            object pO = null;
            object dStartNight = null;
            object iNights = null;
            object objAvail = null;
            object intProdKey = null;
            object bIsTeleBooking = null;
            object objAvailEntry = null;
            object bNoResults = null;
            object bRenderedSummary = null;
            object intIndex = null;
            object intIndexSupplier = null;
            object objFuzzyStayOptions = null;
            object objFuzzyStay = null;
            object bPreciseMatch = null;
            object bStayHasLocalAvail = null;
            object objSuppliersForStay = null;
            object objSupplier = null;
            object objDictAvaiStays = null;
            object strAvailStayKey = null;
            object aryStay = null;
            object sStayNo = null;
            object bStayIndicative = null;
            object bRenderedInitialStay = null;
            object iStayNum = null;
            object ReqDictTemp = null;
            object BookingType = null; /* Undeclared in source */

            pO = _.OBJ(_.CALLm1v0(this, objRenderSettings, "OutputWriter"));
            dStartNight = _.VAL(_.CALLm2v0(this, objRenderSettings, "BookingRequirement", "VisitDate"));
            iNights = _.VAL(_.CALLm2v0(this, objRenderSettings, "BookingRequirement", "Nights"));

            // This is new, VB Polling approach (only supports accommodation, but handles results from
            // multiple providers)
            objAvail = _.OBJ(_.CALLm1v0(this, objData, "Availability"));
            intProdKey = _.VAL(_.CALLm1v0(this, objData, "Product_Key"));
            bIsTeleBooking = _.VAL(_.CALLm1v0(this, objData, "IsOnTeleBookingChannel"));

            objDictAvaiStays = _.OBJ(_.CALLm1v1(this, _outer.Server, "CreateObject", "Scripting.Dictionary"));

            // Quick situation assertion
            if (_.IF(_.NOTEQ(_.NullableSTR(_.CALLm1v0(this, objRenderSettings, "BookingType")), "accommodation")))
            {
                _.RAISEERROR(VBScriptConstants.vbObjectError, "ETWP.BookingUnitSelection", _.CONCAT("BookingUI_StayMain_Polling: BookingType not supported (\"", BookingType, "\")"));
            }
            if (_.IF(_.GT(_.NullableNUM(_.CALLm2v0(this, objRenderSettings, "BookingRequirement", "Offer")), (Int16)0)))
            {
                _.RAISEERROR(VBScriptConstants.vbObjectError, "ETWP.BookingUnitSelection", _.CONCAT("BookingUI_StayMain_Polling: Not supported with Conference Bookings (OfferKey = ", _.CALLm2v0(this, objRenderSettings, "BookingRequirement", "Offer"), ")"));
            }

            // Grab hold of the data for the stay(s) - ensure we've got some availability
            objFuzzyStayOptions = _.OBJ(_.CALLm1argp(this, objAvail, "GetUniqueFuzzyCombinations", _.ARGS.ForceBrackets()));
            if (_.IF(_.EQ(_.NullableNUM(_.CALLm1v0(this, objFuzzyStayOptions, "Count")), (Int16)0)))
            {
                _.CALLm1v1(this, _outer.Page, "PrintTraceWarning", "objAvail.GetUniqueFuzzyCombinations reported zero stay options");
                bNoResults = true;
            }
            else
            {
                // Double-check that all stay options report availability - there shouldn't be any stay
                // data returned that doesn't have avail data

                bNoResults = false;
                var loopEnd10 = _.NUM(_.SUBT(_.CALLm1v0(this, objFuzzyStayOptions, "Count"), (Int16)1));
                var loopStart10 = _.NUM((Int16)0, loopEnd10, (Int16)1);
                if (_.StrictLTE(loopStart10, loopEnd10))
                {
                    for (intIndex = loopStart10; _.StrictLTE(intIndex, loopEnd10); intIndex = _.ADD(intIndex, (Int16)1))
                    {
                        objFuzzyStay = _.OBJ(_.CALLm1argp(this, objFuzzyStayOptions, "GetItem", _.ARGS.Ref(intIndex, v64 => { intIndex = v64; })));
                        objSuppliersForStay = _.OBJ(_.CALLm1v2(this, objAvail, "GetSupplierUnitDataForStay", _.CALLm1v0(this, objFuzzyStay, "StartDate"), _.CALLm1v0(this, objFuzzyStay, "Nights")));
                        if (_.IF(_.EQ(_.NullableNUM(_.CALLm1v0(this, objSuppliersForStay, "Count")), (Int16)0)))
                        {
                            _.CALLm1v1(this, _outer.Page, "PrintTraceWarning", _.CONCAT("Stay (", _.CALLm1v0(this, objFuzzyStay, "StartDate"), ", ", _.CALLm1v0(this, objFuzzyStay, "Nights"), ") reported zero suppliers"));
                            bNoResults = true;
                        }
                        else
                        {
                            var loopEnd11 = _.NUM(_.SUBT(_.CALLm1v0(this, objSuppliersForStay, "Count"), (Int16)1));
                            var loopStart11 = _.NUM((Int16)0, loopEnd11, (Int16)1);
                            if (_.StrictLTE(loopStart11, loopEnd11))
                            {
                                for (intIndexSupplier = loopStart11; _.StrictLTE(intIndexSupplier, loopEnd11); intIndexSupplier = _.ADD(intIndexSupplier, (Int16)1))
                                {
                                    objSupplier = _.OBJ(_.CALLm1argp(this, objSuppliersForStay, "GetItem", _.ARGS.Ref(intIndexSupplier, v65 => { intIndexSupplier = v65; })));
                                    if (_.IF(_.EQ(_.NullableNUM(_.CALLm2v0(this, objSupplier, "Units", "Count")), (Int16)0)))
                                    {
                                        _.CALLm1v1(this, _outer.Page, "PrintTraceWarning", _.CONCAT("Supplier ", _.CALLm1v0(this, objSupplier, "Name"), " for Stay (", _.CALLm1v0(this, objFuzzyStay, "StartDate"), ", ", _.CALLm1v0(this, objFuzzyStay, "Nights"), ") reported zero units"));
                                        bNoResults = true;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // If not, render error and get out
            if (_.IF(bNoResults))
            {
                // Render message, set ProdHasAvail To False (only
                // used by BookingKeys control, I think) and close recordsets
                object byrefalias17 = objRenderSettings;
                try
                {
                    _.CALLm1argp(this, _outer, "RenderNoAvailElement", _.ARGS.Ref(byrefalias17, v66 => { byrefalias17 = v66; }));
                }
                finally { objRenderSettings = byrefalias17; }
                _outer.bProdHasAvail = false; // This is exposed through the WSC's public property "ProdHasAvail"
                return BookingUI_StayMain_Polling_retVal;
            }

            _outer.bProdHasAvail = true; // This is exposed through the WSC's public property "ProdHasAvail"

            // Loop through different stay options
            // - Store data for all stays for calendar
            if (_.IF(_outer.bRenderAsCalendar))
            {

                var loopEnd12 = _.NUM(_.SUBT(_.CALLm1v0(this, objFuzzyStayOptions, "Count"), (Int16)1));
                var loopStart12 = _.NUM((Int16)0, loopEnd12, (Int16)1);
                if (_.StrictLTE(loopStart12, loopEnd12))
                {
                    for (intIndex = loopStart12; _.StrictLTE(intIndex, loopEnd12); intIndex = _.ADD(intIndex, (Int16)1))
                    {
                        objFuzzyStay = _.OBJ(_.CALLm1argp(this, objFuzzyStayOptions, "GetItem", _.ARGS.Ref(intIndex, v67 => { intIndex = v67; })));

                        strAvailStayKey = _.CONCAT("sd_", _.CALLm1v0(this, objFuzzyStay, "StartDate"));
                        if (_.IF(_.CALLm1argp(this, objDictAvaiStays, "Exists", _.ARGS.Ref(strAvailStayKey, v68 => { strAvailStayKey = v68; }))))
                        {

                            // We expect value in the format [stayNo]_[indicative]
                            aryStay = _.SPLIT(_.CALLm0argp(this, objDictAvaiStays, _.ARGS.Ref(strAvailStayKey, v69 => { strAvailStayKey = v69; })), "_");
                            sStayNo = _.VAL(_.CALLm0argp(this, aryStay, _.ARGS.Val((Int16)0)));
                            sStayNo = _.CONCAT(sStayNo, "-", intIndex);

                            bStayIndicative = _.CBOOL(_.CALLm0argp(this, aryStay, _.ARGS.Val((Int16)1)));
                            if (_.IF(_.AND(_.NOT(bStayIndicative), _.CALLm1v0(this, objFuzzyStay, "Indicative"))))
                            {
                                bStayIndicative = _.VAL(_.CALLm1v0(this, objFuzzyStay, "Indicative"));
                            }
                            _.SETm0a1(this, objDictAvaiStays, strAvailStayKey, _.CONCAT(sStayNo, "_", bStayIndicative));
                            _.ERASE(aryStay, v71 => { aryStay = v71; });
                        }
                        else
                        {
                            _.CALLm1v2(this, objDictAvaiStays, "Add", _.CONCAT("sd_", _.CALLm1v0(this, objFuzzyStay, "StartDate")), _.CONCAT(_.ADD(intIndex, (Int16)1), "_", _.CALLm1v0(this, objFuzzyStay, "Indicative")));
                        }

                    }
                }

            }

            bRenderedInitialStay = false;
            // - For unit selections: If we have a perfect match stay, don't bother with the fuzzy options
            var loopEnd13 = _.NUM(_.SUBT(_.CALLm1v0(this, objFuzzyStayOptions, "Count"), (Int16)1));
            var loopStart13 = _.NUM((Int16)0, loopEnd13, (Int16)1);
            if (_.StrictLTE(loopStart13, loopEnd13))
            {
                for (intIndex = loopStart13; _.StrictLTE(intIndex, loopEnd13); intIndex = _.ADD(intIndex, (Int16)1))
                {
                    // 2010-11-03 TB: Stay numbers are 1-based so add 1 to zero-based index
                    iStayNum = _.ADD(intIndex, (Int16)1);

                    objFuzzyStay = _.OBJ(_.CALLm1argp(this, objFuzzyStayOptions, "GetItem", _.ARGS.Ref(intIndex, v72 => { intIndex = v72; })));

                    // 2010-01-29 DWR: Need to use DateValue here since dStartNight might be a string
                    // which will cause the comparison to fail when they represent the same date
                    bPreciseMatch = _.AND(_.EQ(_.DATEVALUE(_.CALLm1v0(this, objFuzzyStay, "StartDate")), _.DATEVALUE(dStartNight)), _.EQ(_.CALLm1v0(this, objFuzzyStay, "Nights"), iNights));

                    // 2010-01-29 DWR: In cases where we have a precise match and we're not rendering a calendar
                    // then we want to just display that stay and get out! If we DON'T have a precise match and
                    // we're not using the calendar approach then we want to render all options and have client
                    // side script juggle them. If we ARE rendering the calendar then we want to display ALL
                    // stays - regardless of whether we have a precise match - because the calendar relies
                    // on the data being in the markup for it to swap around.
                    if (_.IF(_.OR(_outer.bRenderAsCalendar, _.NOT(bPreciseMatch))))
                    {
                        _.CALLm1v1(this, pO, "Write", _.CONCAT("<div class=\"PollingFuzzySetWrapper\" id=\"stay_", iStayNum, "\">"));
                    }

                    // we only render the first stay when initially loading the unitselection
                    // we get the rest via a partial render request and the data is returned as JSON
                    // ready for manipulation and insertion by javascript
                    // this is done to avoid large amounts of HTML being rendered and then hidden
                    if (_.IF(_.NOT(bRenderedInitialStay)))
                    {
                        object byrefalias18 = objRenderSettings;
                        try
                        {
                            _.CALLm1argp(this, _outer, "RenderStay", _.ARGS.Ref(objFuzzyStay, v73 => { objFuzzyStay = v73; }).Ref(objAvail, v74 => { objAvail = v74; }).Ref(iStayNum, v75 => { iStayNum = v75; }).Ref(byrefalias18, v76 => { byrefalias18 = v76; }).Ref(bIsTeleBooking, v77 => { bIsTeleBooking = v77; }).Val(_.CALLm1v0(this, objData, "bookingweb")).Val(_.CALLm1v0(this, objData, "EviivoId")).Val(_.CALLm1v0(this, objData, "Units")));
                        }
                        finally { objRenderSettings = byrefalias18; }
                    }

                    if (_.IF(_outer.bRenderAsCalendar))
                    {
                        bRenderedInitialStay = true;
                    }

                    // Close the wrapper for the current stay date/length result set
                    // 2010-01-29: See earlier comment about this..
                    if (_.IF(_.OR(_outer.bRenderAsCalendar, _.NOT(bPreciseMatch))))
                    {
                        _.CALLm1v1(this, pO, "Write", "</div>");
                    }

                    // If these options were a perfect match, drop out
                    // 2010-01-29 DWR: Unless we're rendering the calendar! In this case client-side javascript
                    // will look after showing one fuzzy stay at a time, but it needs all data present.
                    if (_.IF(_.AND(bPreciseMatch, _.NOT(_outer.bRenderAsCalendar))))
                    {
                        break;
                    }

                }
            }

            if (_.IF(_outer.bRenderAsCalendar))
            {

                ReqDictTemp = _.OBJ(_.CALLm2v1(this, _outer.Page, "Functions", "GetNewObject", "RequestDict"));
                _.CALLm1v2(this, ReqDictTemp, "ForceAdd", "AsyncAction", "unitselection");
                _.CALLm1v2(this, ReqDictTemp, "ForceAdd", "PartialRenderControlList", _.CALLm1v0(this, _outer.Context, "PageControlKey"));
                _.CALLm1v2(this, ReqDictTemp, "ForceAdd", "Silent", "1");
                _.CALLm1v1(this, ReqDictTemp, "Remove", "Debug");
                _.CALLm1v1(this, ReqDictTemp, "Remove", "PartialRenderType");
                _.CALLm1v1(this, ReqDictTemp, "Remove", "Trace");

                _.CALLm1v1(this, _outer.Page, "PrintTrace", "BookingUI_StayMain_Polling: Render available stays as calendars - start");
                _.CALLm1argp(this, _outer, "BookingUI_RenderAvailCal", _.ARGS.Ref(pO, v78 => { pO = v78; }).Ref(objDictAvaiStays, v79 => { objDictAvaiStays = v79; }).Val(false));
                _.CALLm1v1(this, _outer.Page, "PrintTrace", "BookingUI_StayMain_Polling: Render available stays as calendars - end");
                _.CALLm1v1(this, pO, "Write", "<script type=\"text/javascript\">");
                _.CALLm1v1(this, pO, "Write", _.CONCAT("NewMind.ETWP.ControlData[", _.CALLm1v0(this, _outer.Context, "PageControlKey"), "] = { "));
                _.CALLm1v1(this, pO, "Write", "UnitSelPartialRenderLink: '");
                _.CALLm1v1(this, pO, "Write", _.CALLm1v1(this, _.CALLm2v1(this, _outer.Page, "Functions", "GetNewObject", "clsJSON"), "EscapeJSON", _.CONCAT("?", _.CALLm1v0(this, ReqDictTemp, "Querystring"))));
                _.CALLm1v1(this, pO, "Write", "'");
                _.CALLm1v1(this, pO, "Write", " };");
                _.CALLm1v1(this, pO, "Write", "NewMind.ETWP.Booking.InitUnitSel();");
                _.CALLm1v1(this, pO, "Write", _.CONCAT("</", "script>"));

            }

            // Kick off the show / hide script for fuzzy result sets now that we've rendered out all
            // the content rather than waiting for page load - hopefully we can remove some of the
            // flicker that occurs otherwise
            _.CALLm1v1(this, pO, "Write", _.CONCAT("<script type=\"text/javascript\">NewMind.ETWP.Booking.InitPollingUnitSel();</", "script>"));

            return BookingUI_StayMain_Polling_retVal;
        }
        public object RenderNotRequiredDateWarning(ref object pO)
        {
            object RenderNotRequiredDateWarning_retVal = null;
            _.CALLm1v1(this, pO, "Write", _.CONCAT("<p class=\"fuzzyWarning\">", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/notrequireddates", "Sorry, we don't have any availability for the dates you requested. These are the nearest available dates for your room and duration requirements."), "</p>"));
            return RenderNotRequiredDateWarning_retVal;
        }
        public object RenderStay(ref object objFuzzyStay, ref object objAvail, ref object intIndex, ref object objRenderSettings, ref object bIsTeleBooking, object strProductBookingWebIfAny, object strEviivoIdIfAny, object objAllUnits)
        {
            object RenderStay_retVal = null;
            object objSuppliersForStay = null;
            object intProdKey = null;
            object dStartNight = null;
            object iNights = null;
            object pO = null;
            object bStayHasLocalAvail = null;
            object intExtSuppliersShown = null;
            object bRenderedStaySummary = null;
            object intIndexSupplier = null;
            object objSupplier = null;
            object bSkipSupplier = null;
            object strBookingStaySummary = null;
            object bExternalSupplier = null;
            object strSupplierId = null;
            object strSupplierName = null;
            object strSupplierQuality = null;
            object strSupplierLogo = null;
            object strSupplierEviivoName = null;
            object intBookingType = null;
            object bPreciseMatch = null;

            // 2011-08-09 DWR: Expect the BookingRequirement in objRenderSettings to be read-only (since it usually comes from Page.Functions.GetSharedObject),
            // so replace it with an editable version (since some methods in here try to mess about with properties on it)
            _.SETm1a0(this, objRenderSettings, "BookingRequirement", _.OBJ(_.CALLm1v1(this, _outer, "GetEditableBookingRequirement", _.CALLm1v0(this, objRenderSettings, "BookingRequirement"))));

            objSuppliersForStay = _.OBJ(_.CALLm1v2(this, objAvail, "GetSupplierUnitDataForStay", _.CALLm1v0(this, objFuzzyStay, "StartDate"), _.CALLm1v0(this, objFuzzyStay, "Nights")));

            intProdKey = _.VAL(_.CALLm1v0(this, objRenderSettings, "ProductKey"));
            dStartNight = _.VAL(_.CALLm2v0(this, objRenderSettings, "BookingRequirement", "VisitDate"));
            iNights = _.VAL(_.CALLm2v0(this, objRenderSettings, "BookingRequirement", "Nights"));
            pO = _.OBJ(_.CALLm1v0(this, objRenderSettings, "OutputWriter"));

            //just need to set these here as we may be coming in direct from partial render request
            _outer.bRenderAsCalendar = _.VAL(_.CALLm1v0(this, objRenderSettings, "RenderAsCalendar"));
            _outer.IsVBPollingEnabled = _.VAL(_.CALLm1v0(this, objRenderSettings, "IsVBPollingEnabled"));

            // Loop through each supplier and render their units
            // - Suppliers will be ordered NewMind, FrontDesk, Other
            // - If "Booking_ForceExternal" is enabled, FrontDesk is treated as "Other"
            // - There is a limit on the number of "Other" entries to be rendered (if ForceExternal
            //   is enabled, then FrontDesk counts towards this limit)
            // - If ForceExternal is not enabled, FrontDesk will only be rendered if there is no
            //   local availability
            bStayHasLocalAvail = false;
            intExtSuppliersShown = (Int16)0;
            bRenderedStaySummary = false;

            bPreciseMatch = _.AND(_.EQ(_.DATEVALUE(_.CALLm1v0(this, objFuzzyStay, "StartDate")), _.DATEVALUE(dStartNight)), _.EQ(_.CALLm1v0(this, objFuzzyStay, "Nights"), iNights));

            var loopEnd14 = _.NUM(_.SUBT(_.CALLm1v0(this, objSuppliersForStay, "Count"), (Int16)1));
            var loopStart14 = _.NUM((Int16)0, loopEnd14, (Int16)1);
            if (_.StrictLTE(loopStart14, loopEnd14))
            {
                for (intIndexSupplier = loopStart14; _.StrictLTE(intIndexSupplier, loopEnd14); intIndexSupplier = _.ADD(intIndexSupplier, (Int16)1))
                {

                    // Get basic supplier data - count FrontDesk as "Other" if ForceExternal enabled
                    objSupplier = _.OBJ(_.CALLm1argp(this, objSuppliersForStay, "GetItem", _.ARGS.Ref(intIndexSupplier, v80 => { intIndexSupplier = v80; })));
                    if (_.IF(_.CALLm1v0(this, objSupplier, "IsLocal")))
                    {
                        bStayHasLocalAvail = true;
                    }
                    bExternalSupplier = _.VAL(_.OR(_.CALLm1v0(this, objSupplier, "IsExternal"), _.AND(_.CALLm1v0(this, objSupplier, "IsRemote"), _.CALLm2v1(this, _outer.Page, "Site", "Params", "Booking_ForceExternal"))));

                    // Don't render FrontDesk if got local avail for this stay and not enabled ForceExternal
                    bSkipSupplier = _.VAL(_.AND(bStayHasLocalAvail, _.AND(_.CALLm1v0(this, objSupplier, "IsRemote"), _.NOT(_.CALLm2v1(this, _outer.Page, "Site", "Params", "Booking_ForceExternal")))));
                    if (_.IF(_.NOT(bSkipSupplier)))
                    {

                        // Don't bother rendering stay summary title if we've got a perfect match, as we
                        // won't be showing any fuzzy content if there's a spot-on option
                        if (_.IF(_.AND(_.NOT(bPreciseMatch), _.NOT(bRenderedStaySummary))))
                        {
                            if (_.IF(_.NOT(_outer.bRenderAsCalendar)))
                            {
                                _.CALLm1argp(this, _outer, "BookingUI_StaySummary", _.ARGS.Ref(dStartNight, v81 => { dStartNight = v81; }).Ref(iNights, v82 => { iNights = v82; }).Val(_.CALLm1v0(this, objFuzzyStay, "StartDate")).Val(_.CALLm1v0(this, objFuzzyStay, "Nights")).Ref(pO, v83 => { pO = v83; }));
                            }
                            bRenderedStaySummary = true;
                        }

                        // If this is an external supplier, we need the deep-link quality to pass to get
                        // included in the hidden booking-info form fields

                        // PW 2010-07-28 I have added a new field called strSupplierEviivoName
                        // This is to pass through the original name field from Eviivo through to the polling exit page.
                        // Previously, we did some manipulation on this value to ensure it had a nice display name.
                        // However, this had broken Eviivo's own external link - we have in the past asked Eviivo to provide a
                        // nice display name field but until they do so we are going to have to do our own and pass both values through as hidden
                        // form fields
                        if (_.IF(bExternalSupplier))
                        {
                            strSupplierId = _.VAL(_.CALLm1v0(this, objSupplier, "ID"));
                            strSupplierName = _.VAL(_.CALLm1v0(this, objSupplier, "DisplayName"));
                            strSupplierQuality = _.VAL(_.CALLm1v0(this, objSupplier, "Quality"));
                            strSupplierEviivoName = _.VAL(_.CALLm1v0(this, objSupplier, "Name"));
                        }
                        else
                        {
                            strSupplierId = VBScriptConstants.Null;
                            strSupplierName = VBScriptConstants.Null;
                            strSupplierQuality = VBScriptConstants.Null;
                            strSupplierEviivoName = VBScriptConstants.Null;
                        }

                        // Render the actual options (wrap in the standard form tag)
                        if (_.IF(_.CALLm1v0(this, objSupplier, "IsLocal")))
                        {
                            if (_.IF(_.ISEMPTY(_outer.IsExternalBooking)))
                            {
                                _.CALLm1argp(this, _outer, "InitExternalBookingSettings", _.ARGS.ForceBrackets());
                            }
                            if (_.IF(_outer.IsExternalBooking))
                            {
                                intBookingType = _.VAL(_outer.BOOKING_Redirect);
                                _outer.strProductEstateID = _.VAL(_.CALLm1argp(this, _outer.DMS, "GetProductEstateID", _.ARGS.Ref(intProdKey, v84 => { intProdKey = v84; })));
                                _outer.strExtBookUrl = _.VAL(_.CALLm1argp(this, _outer, "GetExtBookUrlFromProductEstate", _.ARGS.Ref(_outer.strProductEstateID, v85 => { _outer.strProductEstateID = v85; })));
                            }
                            else
                            {
                                intBookingType = _.VAL(_outer.BOOKING_Local);
                            }
                        }
                        else if (_.IF(_.CALLm1v0(this, objSupplier, "IsExternal")))
                        {
                            // 2011-07-20 DWR: We don't need to call InitExternalBookingSettings if dealing with an VB Polling product as
                            // the next page should always be the Polling Exit (no point redirecting to another site which will then - if
                            // it's an NM site - have to display another redirect page to book the product)
                            intBookingType = _.VAL(_outer.BOOKING_PollingRedirect);
                        }
                        else
                        {
                            if (_.IF(_.ISEMPTY(_outer.IsExternalBooking)))
                            {
                                _.CALLm1argp(this, _outer, "InitExternalBookingSettings", _.ARGS.ForceBrackets());
                            }
                            if (_.IF(_outer.IsExternalBooking))
                            {
                                intBookingType = _.VAL(_outer.BOOKING_Redirect);
                                _outer.strProductEstateID = _.VAL(_.CALLm1argp(this, _outer.DMS, "GetProductEstateID", _.ARGS.Ref(intProdKey, v86 => { intProdKey = v86; })));
                                _outer.strExtBookUrl = _.VAL(_.CALLm1argp(this, _outer, "GetExtBookUrlFromProductEstate", _.ARGS.Ref(_outer.strProductEstateID, v87 => { _outer.strProductEstateID = v87; })));
                            }
                            else
                            {
                                intBookingType = _.VAL(_outer.BOOKING_Eviivo);
                            }
                        }

                        // Local and FrontDesk both use current site name w/out logo
                        // External Suppliers should have their own logo passed in
                        // PW - 	moved this out of BookingUI_StayDetails_PollingHeader
                        //		we can now pass it to the hidden form fields
                        //		for use on the polling exit page
                        if (_.IF(_.CALLm1v0(this, objSupplier, "IsExternal")))
                        {
                            strSupplierLogo = _.VAL(_.CALLm1v0(this, objSupplier, "Logo"));
                            if (_.IF(_.EQ(_.NullableSTR(_.TRIM(_.CONCAT("", strSupplierName))), "")))
                            {
                                strSupplierName = "Unnamed Supplier";
                            }
                            else if (_.IF(_.EQ(_.NullableSTR(strSupplierLogo), "")))
                            {
                                // 2014-07-01 DWR: It's common for Eviivo to not return logo data for the Polling Providers so for most cases we take the Supplier Name (the
                                // Eviivo version, rather than the "friendly" version that we maintain) and request the logo from ntop using it. For cases where Eviivo
                                // results are treated as Polling results (see FogBugz 10386), we need a special case (the friendly name will always be "Eviivo" in
                                // this case).
                                if (_.IF(_.EQ(_.NullableSTR(strSupplierName), "Eviivo")))
                                {
                                    strSupplierLogo = _.VAL(_.CALLm1v2(this, _outer.Page, "ImageResource", "bookonline/unitselection/polling/eviivo", "/engine/shared_gfx/eviiopollingresult.jpg"));
                                }
                                else
                                {
                                    // 2008-12-09 DWR: Supplier Logo isn't actually going to be received from the Eviivo Component, we mash Supplier Name into this url
                                    // 2010-03-04 DWR: Eviivo moved the logo location..
                                    strSupplierLogo = _.CONCAT("http://www.ntopsearch.com/media/images/Suppliers/", strSupplierEviivoName, ".gif");
                                }
                            }
                        }
                        else
                        {
                            // 2009-02-12 DWR: Changed the way in which supplier name and logo are determined for Local / FrontDesk
                            // suppliers (ie. the non-external entries) - before it had no logo and displayed the site name, now
                            // these are the defaults, but content can be pulled from languages xml. This content can be specified
                            // to vary per estate if desired (intended to be used when VB Polling is combined with Force External
                            // Bookings)
                            // - Supplier Name
                            strSupplierName = "";
                            if (_.IF(_outer.IsExternalBooking))
                            {
                                strSupplierName = _.VAL(_.CALLm1v2(this, _outer.Page, "Resource", _.CONCAT("bookonline/unitselection/polling/localsupplier/estate_", _outer.strProductEstateID, "/name"), ""));
                            }
                            if (_.IF(_.EQ(_.NullableSTR(strSupplierName), "")))
                            {
                                //#MJ -	the resource manage is the same for both main sites and channel sites
                                //		therefore we can never use Page.Site.Name as an alternative value as this would be cached wrongly by the ResourceManager
                                //		so try to pull one from there, if not fall back to the site name
                                strSupplierName = _.VAL(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/polling/localsupplier/name", ""));
                                if (_.IF(_.EQ(_.NullableSTR(strSupplierName), "")))
                                {
                                    strSupplierName = _.VAL(_.CALLm2v0(this, _outer.Page, "Site", "Name"));
                                }
                            }
                            // - Supplier Logo
                            strSupplierLogo = _.VAL(_.CALLm1argp(this, _outer, "GetSupplierLogo", _.ARGS.Ref(_outer.strProductEstateID, v88 => { _outer.strProductEstateID = v88; })));

                        }

                        // 2013-02-05 TB: objRenderSettings is used by RenderBookingInfoForm to populate some hidden stay information
                        // For fuzzy stays, both nights and startdate may differ from the original requirements.
                        // For FogBugz case 7594 I added the second line below which wasn't present.
                        _.SETm1a0(this, _.CALLm1v0(this, objRenderSettings, "BookingRequirement"), "VisitDate", _.VAL(_.CALLm1v0(this, objFuzzyStay, "StartDate")));
                        _.SETm1a0(this, _.CALLm1v0(this, objRenderSettings, "BookingRequirement"), "Nights", _.VAL(_.CALLm1v0(this, objFuzzyStay, "Nights")));

                        // 2014-03-12 DWR: We need to pass the Search Industry Classification into the form rendering code for VB Polling Products so that the
                        // Polling Exist can generate the deep link correctly. An Eviivo Configset can be set up with zero, meaning support either 1 OR 9. The
                        // Avail Component will perform searches for both in that case but only allow any Products to return results for one. Since we won't
                        // get an objSupplier reference with zero units (since that would mean it's not got availability and we're only looking at available
                        // options here) we can just grab the IndustryClassification values from the first Unit since it is guaranteed to be consistent
                        // across all Units for this booking option. The IndustryClassification value will be zero for non-Eviivo data but that won't
                        // matter since it's only ever consider in the Polling Exit which is for Eviivo results only.
                        object byrefalias19 = objRenderSettings;
                        try
                        {
                            _.CALLm1argp(this, _outer, "RenderBookingInfoForm", _.ARGS.Ref(pO, v89 => { pO = v89; }).Ref(intProdKey, v90 => { intProdKey = v90; }).Ref(byrefalias19, v91 => { byrefalias19 = v91; }).Ref(intBookingType, v92 => { intBookingType = v92; }).Ref(strSupplierId, v93 => { strSupplierId = v93; }).Ref(strSupplierName, v94 => { strSupplierName = v94; }).Ref(strSupplierEviivoName, v95 => { strSupplierEviivoName = v95; }).Ref(strSupplierQuality, v96 => { strSupplierQuality = v96; }).Ref(strSupplierLogo, v97 => { strSupplierLogo = v97; }).Val(_.CALLm1v0(this, _.CALLm2v1(this, objSupplier, "Units", "GetItem", (Int16)0), "IndustryClassification")));
                        }
                        finally { objRenderSettings = byrefalias19; }

                        _.CALLm1argp(this, _outer, "BookingUI_StayDetails_PollingHeader", _.ARGS.Ref(objSupplier, v98 => { objSupplier = v98; }).Ref(pO, v99 => { pO = v99; }).Ref(strSupplierLogo, v100 => { strSupplierLogo = v100; }).Ref(strSupplierName, v101 => { strSupplierName = v101; }));

                        // 2009-09-14 DWR: Forcing iStayNum to "1" every time - since we are clearly only having
                        // one stay per form (since we open the form above - in RenderBookingInfoForm - and we
                        // close it below) we'll always be passing only a single stay to the next stage. This
                        // makes things easier - the multiple-stays-per-form idea was ridiculous.
                        // 2010-10-21 TB: Changing back to use unique stay index. Multiple stays per form will
                        // happen for fuzzy results and calendar view. html ids use the stay key, as does the JS
                        // when choosing to show/hide the book now button.
                        object byrefalias20 = intIndex, byrefalias21 = bIsTeleBooking;
                        try
                        {
                            _.CALLm1argp(this, _outer, "BookingUI_StayDetails", _.ARGS.Ref(objSupplier, v102 => { objSupplier = v102; }).Ref(byrefalias20, v103 => { byrefalias20 = v103; }).Ref(dStartNight, v104 => { dStartNight = v104; }).Ref(iNights, v105 => { iNights = v105; }).Ref(byrefalias21, v106 => { byrefalias21 = v106; }).Ref(strProductBookingWebIfAny, v107 => { strProductBookingWebIfAny = v107; }).Ref(strEviivoIdIfAny, v108 => { strEviivoIdIfAny = v108; }).Val(_.CALLm1v0(this, objRenderSettings, "ProductKey")).Val(_.CALLm1v0(this, objRenderSettings, "Channel")).Val(_.CALLm1v0(this, objFuzzyStay, "Indicative")).Val(_.NOT(_.CALLm1v0(this, objFuzzyStay, "HasInvalidIndicative"))).Ref(objAllUnits, v109 => { objAllUnits = v109; }).Val(VBScriptConstants.Nothing).Ref(pO, v110 => { pO = v110; }).Val(false));
                        }
                        finally { intIndex = byrefalias20; bIsTeleBooking = byrefalias21; } //we can't render the maximum available units for polling

                        _.CALLm1v1(this, pO, "Write", "</form>");

                    }

                }
            }

            return RenderStay_retVal;
        }
        public object RenderNoAvailElement(object objRenderSettings)
        {
            object RenderNoAvailElement_retVal = null;
            object pO = null;
            object strClassMonth = null;

            pO = _.OBJ(_.CALLm1v0(this, objRenderSettings, "OutputWriter"));
            _.CALLm1v1(this, pO, "Write", "<div class=\"pnNoAvail\">");
            _.CALLm1v1(this, pO, "Write", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/noavailability", "<p>No availability for this product for the specified date. This may occur if the accommodation is booked prior to your arrival at this page.</p>"));
            _.CALLm1v1(this, pO, "Write", "</div>");

            if (_.IF(_.CALLm1v0(this, objRenderSettings, "RenderAsCalendar")))
            {

                strClassMonth = "MonthWrapper";

                _.CALLm1v1(this, pO, "Write", "<div class=\"CalendarsWrapper\">");
                //
                _.CALLm1argp(this, _outer, "BookingUI_RenderCalendarMonth", _.ARGS.Ref(pO, v111 => { pO = v111; }).Val(_.CALLm2v0(this, objRenderSettings, "BookingRequirement", "VisitDate")).Val(_.CONCAT(strClassMonth, " currentmonth")));
                //					' last day + 1 to get the first day of the next month for the calendar
                _.CALLm1argp(this, _outer, "BookingUI_RenderCalendarMonth", _.ARGS.Ref(pO, v112 => { pO = v112; }).Val(_.ADD(_.CALLm3v1(this, _outer.Page, "Functions", "Dates", "fn_GetLastDateOfMonth", _.CALLm2v0(this, objRenderSettings, "BookingRequirement", "VisitDate")), (Int16)1)).Val(_.CONCAT(strClassMonth, " nextmonth")));

                // global count used to track how many calendars have been added to the output for the prev/next buttons
                _outer.g_iNumberOfCalendarsRendered = (Int16)2;

                _.CALLm1argp(this, _outer, "BookingUI_RenderAvailCalLinks", _.ARGS.Val(_.CALLm2v0(this, objRenderSettings, "BookingRequirement", "VisitDate")).Ref(pO, v113 => { pO = v113; }));
                _.CALLm1argp(this, _outer, "BookingUI_RenderAvailCalKey", _.ARGS.Ref(pO, v114 => { pO = v114; }));
                _.CALLm1v1(this, pO, "Write", "</div>");

                _.CALLm1v1(this, pO, "Write", _.CONCAT("<script type=\"text/javascript\">NewMind.ETWP.Booking.UpdateCalLinks();</", "script>"));
            }

            return RenderNoAvailElement_retVal;
        }
        // ====================================================================================================
        // RENDER: Main entry point when VB Polling is disabled (or handling tickets, not acco products)
        // ====================================================================================================
        public object BookingUI_StayMain_Legacy(ref object objData, ref object objRenderSettings)
        {
            object BookingUI_StayMain_Legacy_retVal = null;
            object pO = null;
            object intBookingType = null;
            object bNoResults = null;
            object objFuzzyStayOptions = null;
            object objFuzzyStay = null;
            object objSuppliersForStay = null;
            object objAvailEntry = null;
            object lsRemoteUnitSelections = null;
            object objAvail = null;
            object intProdKey = null;
            object bIsTeleBooking = null;
            // This is the non-VB-Polling approach (supports EITHER FrontDesk OR local availability for accommodation)
            //reset the output variable to our OutputWriter
            pO = _.OBJ(_.CALLm1v0(this, objRenderSettings, "OutputWriter"));

            objAvail = _.OBJ(_.CALLm1v0(this, objData, "Availability"));
            intProdKey = _.VAL(_.CALLm1v0(this, objData, "Product_Key"));

            bIsTeleBooking = _.VAL(_.CALLm1v0(this, objData, "IsOnTeleBookingChannel"));

            // Grab hold of the data (in this method, there should only ever be zero or one fuzzy
            // stay options, as the BookingUI_StayMain_Legacy method handle fuzzy availability)
            objFuzzyStayOptions = _.OBJ(_.CALLm1argp(this, objAvail, "GetUniqueFuzzyCombinations", _.ARGS.ForceBrackets()));
            if (_.IF(_.EQ(_.NullableNUM(_.CALLm1v0(this, objFuzzyStayOptions, "Count")), (Int16)0)))
            {
                _.CALLm1v1(this, _outer.Page, "PrintTraceWarning", "objAvail.GetUniqueFuzzyCombinations reported zero stay options");
                bNoResults = true;
            }
            else
            {
                // Any suppliers returned here will be sorted with Local / NewMind first, then FrontDesk
                // second (if we have both) - if there are multiple, it should always be the first one
                // that we want
                objFuzzyStay = _.OBJ(_.CALLm1v1(this, objFuzzyStayOptions, "GetItem", (Int16)0));
                _.CALLm1v1(this, _outer.Page, "PrintTrace", _.CONCAT("BookingUI_StayMain_Legacy: Get data for stay - ", _.CALLm1v0(this, objFuzzyStay, "StartDate"), ", ", _.CALLm1v0(this, objFuzzyStay, "Nights")));
                objSuppliersForStay = _.OBJ(_.CALLm1v2(this, objAvail, "GetSupplierUnitDataForStay", _.CALLm1v0(this, objFuzzyStay, "StartDate"), _.CALLm1v0(this, objFuzzyStay, "Nights")));
                if (_.IF(_.EQ(_.NullableNUM(_.CALLm1v0(this, objSuppliersForStay, "Count")), (Int16)0)))
                {
                    _.CALLm1v1(this, _outer.Page, "PrintTraceWarning", "objAvail.GetSupplierUnitDataForStay reported zero suppliers");
                    bNoResults = true;
                }
                else
                {
                    objAvailEntry = _.OBJ(_.CALLm1v1(this, objSuppliersForStay, "GetItem", (Int16)0));
                    bNoResults = false;
                }
            }

            // Open form and prepare to wrap content in "staySelection" container
            if (_.IF(_outer.IsExternalBooking))
            {
                intBookingType = _.VAL(_outer.BOOKING_Redirect);
            }
            else
            {
                intBookingType = _.VAL(_outer.BOOKING_Local);
            }

            object byrefalias22 = objRenderSettings;
            try
            {
                _.CALLm1argp(this, _outer, "RenderBookingInfoForm", _.ARGS.Ref(pO, v115 => { pO = v115; }).Ref(intProdKey, v116 => { intProdKey = v116; }).Ref(byrefalias22, v117 => { byrefalias22 = v117; }).Ref(intBookingType, v118 => { intBookingType = v118; }).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null));
            }
            finally { objRenderSettings = byrefalias22; }

            _.CALLm1v1(this, pO, "Write", "<div class=\"staySelection\">");

            // Render info (or display warning if no availability)
            if (_.IF(bNoResults))
            {
                object byrefalias23 = objRenderSettings;
                try
                {
                    _.CALLm1argp(this, _outer, "RenderNoAvailElement", _.ARGS.Ref(byrefalias23, v119 => { byrefalias23 = v119; }));
                }
                finally { objRenderSettings = byrefalias23; }
                _outer.bProdHasAvail = false; // This is exposed through the WSC's public property "ProdHasAvail"
            }
            else
            {
                _outer.bProdHasAvail = true; // This is exposed through the WSC's public property "ProdHasAvail"
                if (_.IF(_.EQ(_.NullableSTR(_.CALLm1v0(this, objRenderSettings, "BookingType")), "accommodation")))
                {

                    // Retrieve any unit selections that have been passed in through the querystring
                    // - eg. when VisitBritain hooks in to complete a booking
                    // There will be an entry in lsUnitSelections for each requirement.
                    // Note that ReqNo in the avail data is one-based while the lsUnitSelections indices
                    // are zero-based, so the UnitKey for ReqNo 1 = lsUnitSelections(0). If there was no
                    // selection made for a ReqNo, the lsUnitSelections value will be zero.
                    // NB: This value might be Nothing if no selections are passed in on querystring.
                    lsRemoteUnitSelections = _.OBJ(_.CALLm1argp(this, _outer, "BookingUI_UnitSel_GetOptionsRemoteSelected", _.ARGS.Ref(objAvailEntry, v120 => { objAvailEntry = v120; })));

                    // Render the unit selection options (pass "1" as iStayNum parameter - we'll only
                    // be rendering a single stay option here, since fuzzy isn't supported in this
                    // configuration..)
                    _.CALLm1argp(this, _outer, "BookingUI_StayDetails", _.ARGS.Ref(objAvailEntry, v121 => { objAvailEntry = v121; }).Val((Int16)1).Val(_.CALLm2v0(this, objRenderSettings, "BookingRequirement", "VisitDate")).Val(_.CALLm2v0(this, objRenderSettings, "BookingRequirement", "Nights")).Ref(bIsTeleBooking, v122 => { bIsTeleBooking = v122; }).Val(_.CALLm1v0(this, objData, "bookingweb")).Val(_.CALLm1v0(this, objData, "EviivoId")).Ref(intProdKey, v123 => { intProdKey = v123; }).Val(_.CALLm1v0(this, objRenderSettings, "Channel")).Val(_.CALLm1v0(this, objFuzzyStay, "Indicative")).Val(_.NOT(_.CALLm1v0(this, objFuzzyStay, "HasInvalidIndicative"))).Val(_.CALLm1v0(this, objData, "Units")).Ref(lsRemoteUnitSelections, v124 => { lsRemoteUnitSelections = v124; }).Ref(pO, v125 => { pO = v125; }).Val(_.CALLm1v0(this, objRenderSettings, "RenderMaximumUnitsAvailable")));

                }
                else
                {
                    _.CALLm1argp(this, _outer, "BookingUI_TicketsSummary", _.ARGS.Ref(objAvailEntry, v126 => { objAvailEntry = v126; }).Val(_.CALLm2v0(this, objRenderSettings, "BookingRequirement", "VisitDate")).Ref(pO, v127 => { pO = v127; }));
                }
            }

            // Close "staySelection" div and form
            _.CALLm1v1(this, pO, "Write", "</div>");
            _.CALLm1v1(this, pO, "Write", "</form>");
            return BookingUI_StayMain_Legacy_retVal;
        }
        // SUMMARY: prepare a list of UnitKey selections for each ReqNo is availability recordset
        // [rsAvail]: ADO unit recordset from availability object
        // <retval>: clsList with as many values as there are ReqNo entries, containing the UnitKey for each one
        public object BookingUI_UnitSel_GetOptionsRemoteSelected(object objAvailEntry)
        {
            object BookingUI_UnitSel_GetOptionsRemoteSelected_retVal = null;
            object intIndex = null;
            object objUnit = null;
            object arrReqUnitOptions = null;
            object arrReqUnitSelections = null;
            object intUnitSel = null;
            object lsUnitKeys = null;
            object BookingUI_UnitSel_GetOptionSelected = null;

            // Build up a list of unit options:
            // - Will get a list of objects where each object has properties:
            //    > ReqNo (integer)
            //    > NumPeople (integer)
            //    > Units (list of integers)
            // - We're going to loop through the availability recordset, so must remember
            //   to return it back to the beginning when we're done
            arrReqUnitOptions = _.OBJ(_.CALLm2v1(this, _outer.Page, "Functions", "GetNewObject", "clsList"));
            var loopEnd15 = _.NUM(_.SUBT(_.CALLm2v0(this, objAvailEntry, "Units", "Count"), (Int16)1));
            var loopStart15 = _.NUM((Int16)0, loopEnd15, (Int16)1);
            if (_.StrictLTE(loopStart15, loopEnd15))
            {
                for (intIndex = loopStart15; _.StrictLTE(intIndex, loopEnd15); intIndex = _.ADD(intIndex, (Int16)1))
                {
                    objUnit = _.OBJ(_.CALLm2argp(this, objAvailEntry, "Units", "GetItem", _.ARGS.Ref(intIndex, v128 => { intIndex = v128; })));
                    _.CALLm1argp(this, _outer, "BookingUI_UnitSel_AddReqUnitOption", _.ARGS.Ref(arrReqUnitOptions, v129 => { arrReqUnitOptions = v129; }).Val(_.CALLm1v0(this, objUnit, "ReqNo")).Val(_.CALLm1v0(this, objUnit, "ReqSize")).Val(_.CALLm1v0(this, objUnit, "UnitKey")));
                    //BookingUI_UnitSel_AddReqUnitOption arrReqUnitOptions, objUnit.ReqNo, objUnit.UnitCount, objUnit.UnitKey
                }
            }

            // Build up a list of unit selections passed in from external site (eg. VisitBritain):
            // - Will get a list of objects where each object has properties:
            //    > NumPeople (integer)
            //    > UnitKey (integer)
            //    > PossReqNos (list of integers)
            //       = list of ReqNo values that this may be
            //         a user selection for
            intUnitSel = (Int16)0;
            arrReqUnitSelections = _.OBJ(_.CALLm2v1(this, _outer.Page, "Functions", "GetNewObject", "clsList"));
            while (true)
            {
                intUnitSel = _.ADD(intUnitSel, (Int16)1);
                if (_.IF(_.GT(_.NullableNUM(_.LEN(_.CALLm0argp(this, _outer.Request, _.ARGS.Val(_.CONCAT("URslt", intUnitSel))))), (Int16)0)))
                {
                    _.CALLm1argp(this, _outer, "BookingUI_UnitSel_AddReqUnitSelection", _.ARGS.Ref(arrReqUnitSelections, v130 => { arrReqUnitSelections = v130; }).RefIfArray(_outer.Request, _.ARGS.Val(_.CONCAT("URslt", intUnitSel))).Ref(arrReqUnitOptions, v131 => { arrReqUnitOptions = v131; }));
                }
                else
                {
                    break;
                }
            }

            // If there were no selections passed in like this, return Nothing
            if (_.IF(_.EQ(_.NullableNUM(_.CALLm1v0(this, arrReqUnitSelections, "Count")), (Int16)0)))
            {
                BookingUI_UnitSel_GetOptionSelected = VBScriptConstants.Nothing;
            }

            // Now try to return matched unit options / selections
            // - Get back a list of unit keys, one key per requirement
            //   (If failed to get a perfect match, some of these values may be zero)
            BookingUI_UnitSel_GetOptionsRemoteSelected_retVal = _.OBJ(_.CALLm1argp(this, _outer, "BookingUI_UnitSel_GetMatchedReqUnitSelection", _.ARGS.Ref(arrReqUnitOptions, v132 => { arrReqUnitOptions = v132; }).Ref(arrReqUnitSelections, v133 => { arrReqUnitSelections = v133; })));

            return BookingUI_UnitSel_GetOptionsRemoteSelected_retVal;
        }
        public object BookingUI_UnitSel_AddReqUnitOption(ref object arrReqUnitOptions, ref object intReqNo, ref object intNumPeople, ref object intUnitKey)
        {
            object BookingUI_UnitSel_AddReqUnitOption_retVal = null;
            object objEntry = null;
            object objEntryPrev = null;

            // Input list SHOULD be initialised as an empty list, but just in case..
            if (_.IF(_.OR(_.ISEMPTY(arrReqUnitOptions), _.ISNULL(arrReqUnitOptions))))
            {
                arrReqUnitOptions = _.OBJ(_.CALLm2v1(this, _outer.Page, "Functions", "GetNewObject", "clsList"));
            }

            // If we've already got list items, check whether we're still working on the same
            // ReqNo as the previous entry. If so, add to that entry's unit list.
            if (_.IF(_.GT(_.NullableNUM(_.CALLm1v0(this, arrReqUnitOptions, "Count")), (Int16)0)))
            {
                objEntryPrev = _.OBJ(_.CALLm0argp(this, arrReqUnitOptions, _.ARGS.Val(_.SUBT(_.CALLm1v0(this, arrReqUnitOptions, "Count"), (Int16)1))));
                if (_.IF(_.EQ(_.CALLm0argp(this, objEntryPrev, _.ARGS.Val("ReqNo")), intReqNo)))
                {
                    object byrefalias24 = intUnitKey;
                    try
                    {
                        _.CALLm1argp(this, _.CALLm0argp(this, objEntryPrev, _.ARGS.Val("Units")), "Add", _.ARGS.Ref(byrefalias24, v134 => { byrefalias24 = v134; }));
                    }
                    finally { intUnitKey = byrefalias24; }
                    return BookingUI_UnitSel_AddReqUnitOption_retVal;
                }
            }

            // Need to create a new entry
            objEntry = _.OBJ(_.CALLm2v1(this, _outer.Page, "Functions", "GetNewObject", "clsValueBag"));
            _.SETm0a1(this, objEntry, "ReqNo", _.VAL(intReqNo));
            _.SETm0a1(this, objEntry, "NumPeople", _.VAL(intNumPeople));
            _.SETm0a1(this, objEntry, "Units", _.OBJ(_.CALLm2v1(this, _outer.Page, "Functions", "GetNewObject", "clsList")));
            object byrefalias25 = intUnitKey;
            try
            {
                _.CALLm1argp(this, _.CALLm0argp(this, objEntry, _.ARGS.Val("Units")), "Add", _.ARGS.Ref(byrefalias25, v135 => { byrefalias25 = v135; }));
            }
            finally { intUnitKey = byrefalias25; }
            _.CALLm1argp(this, arrReqUnitOptions, "Add", _.ARGS.Ref(objEntry, v136 => { objEntry = v136; }));

            return BookingUI_UnitSel_AddReqUnitOption_retVal;
        }
        public object BookingUI_UnitSel_AddReqUnitSelection(ref object arrReqUnitSelections, ref object strUnitSelInfo, ref object arrReqUnitOptions)
        {
            object BookingUI_UnitSel_AddReqUnitSelection_retVal = null;
            int errOn = _.GETERRORTRAPPINGTOKEN();
            object arrSegments = null;
            object intNumAdults = null;
            object intNumChildren = null;
            object intUnitKey = null;
            object intIndex = null;
            object objEntry = null;
            object objUnitList = null;

            // Input list SHOULD be initialised as an empty list, but just in case..
            bool ifResult;
            object byrefalias26 = arrReqUnitSelections;
            try
            {
                ifResult = _.IF(() => _.OR(_.ISEMPTY(byrefalias26), _.ISNULL(byrefalias26)), errOn);
            }
            finally { arrReqUnitSelections = byrefalias26; }
            if (ifResult)
            {
                object byrefalias27 = arrReqUnitSelections;
                try
                {
                    byrefalias27 = _.OBJ(_.CALLm2v1(this, _outer.Page, "Functions", "GetNewObject", "clsList"));
                }
                finally { arrReqUnitSelections = byrefalias27; }
            }

            // strUnitSelInfo should be of the form "UnitKey,NumAdults,NumChildren"
            // Exit if not
            object byrefalias28 = strUnitSelInfo;
            try
            {
                arrSegments = _.SPLIT(byrefalias28, ",");
            }
            finally { strUnitSelInfo = byrefalias28; }
            if (_.IF(_.NOTEQ(_.NullableNUM(_.UBOUND(arrSegments)), (Int16)2)))
            {
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return BookingUI_UnitSel_AddReqUnitSelection_retVal;
            }

            // Ensure entries in string are numeric (exit if not)
            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            _.HANDLEERROR(errOn, () => {
                intUnitKey = _.CLNG(_.CALLm0argp(this, arrSegments, _.ARGS.Val((Int16)0)));
            });
            if (_.IF(() => _.ERR, errOn))
            {
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return BookingUI_UnitSel_AddReqUnitSelection_retVal;
            }
            _.HANDLEERROR(errOn, () => {
                intNumAdults = _.CLNG(_.CALLm0argp(this, arrSegments, _.ARGS.Val((Int16)1)));
            });
            if (_.IF(() => _.ERR, errOn))
            {
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return BookingUI_UnitSel_AddReqUnitSelection_retVal;
            }
            _.HANDLEERROR(errOn, () => {
                intNumChildren = _.CLNG(_.CALLm0argp(this, arrSegments, _.ARGS.Val((Int16)2)));
            });
            if (_.IF(() => _.ERR, errOn))
            {
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return BookingUI_UnitSel_AddReqUnitSelection_retVal;
            }
            _.STOPERRORTRAPPINGANDCLEARANYERROR(errOn);

            // Ensure values look reasonable
            if (_.IF(_.OR(_.OR(_.OR(_.LTE(_.NullableNUM(intUnitKey), (Int16)0), _.LT(_.NullableNUM(intNumAdults), (Int16)0)), _.LT(_.NullableNUM(intNumChildren), (Int16)0)), _.LTE(_.NullableNUM(_.ADD(intNumAdults, intNumChildren)), (Int16)0))))
            {
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return BookingUI_UnitSel_AddReqUnitSelection_retVal;
            }

            // Preparer new entry
            objEntry = _.OBJ(_.CALLm2v1(this, _outer.Page, "Functions", "GetNewObject", "clsValueBag"));
            _.SETm0a1(this, objEntry, "NumPeople", _.ADD(intNumAdults, intNumChildren));
            _.SETm0a1(this, objEntry, "UnitKey", _.VAL(intUnitKey));
            _.SETm0a1(this, objEntry, "PossReqNos", _.OBJ(_.CALLm2v1(this, _outer.Page, "Functions", "GetNewObject", "clsList")));

            // Look through the unit options and look for possible requirement matches
            // - We've got a set of requirement / room options from the DMS and we've (possibly) got a
            //   set of unit selections from VisitBritain (or whoever), but these may not currently be
            //   aligned, so we want to determine the possible ways they MIGHT go together, and we'll
            //   try to get the best configuration (which will hopefully match the original choice)
            //   later on.
            bool ifResult2;
            object byrefalias29 = arrReqUnitOptions;
            ifResult2 = _.IF(() => _.GT(_.NullableNUM(_.CALLm1v0(this, byrefalias29, "Count")), (Int16)0), errOn);
            if (ifResult2)
            {
                object loopEnd16 = 0, loopStart16 = 0;
                var loopConstraintsInitialized = false;
                object byrefalias30 = arrReqUnitOptions;
                _.HANDLEERROR(errOn, () => {
                    loopEnd16 = _.NUM(_.SUBT(_.CALLm1v0(this, byrefalias30, "Count"), (Int16)1));
                    loopStart16 = _.NUM((Int16)0);
                    if ((loopStart16 is DateTime) || (loopStart16 is Decimal))
                        intIndex = loopStart16;
                    loopStart16 = _.NUM((Int16)0, loopEnd16, (Int16)1);
                    loopConstraintsInitialized = true;
                });
                if (_.StrictLTE(loopStart16, loopEnd16))
                {
                    if (loopConstraintsInitialized)
                        intIndex = loopStart16;
                    while (true)
                    {
                        // If requirement option matches the selection's NumPeople and contains the
                        // UnitKey, then we've got a possible match
                        bool ifResult3;
                        object byrefalias31 = arrReqUnitOptions;
                        ifResult3 = _.IF(() => _.AND(_.EQ(_.CALLm0argp(this, _.CALLm0argp(this, byrefalias31, _.ARGS.Ref(intIndex, v139 => { intIndex = v139; })), _.ARGS.Val("NumPeople")), _.CALLm0argp(this, objEntry, _.ARGS.Val("NumPeople"))), _.CALLm1argp(this, _.CALLm0argp(this, _.CALLm0argp(this, byrefalias31, _.ARGS.Ref(intIndex, v140 => { intIndex = v140; })), _.ARGS.Val("Units")), "Contains", _.ARGS.RefIfArray(objEntry, _.ARGS.Val("UnitKey")))), errOn);
                        if (ifResult3)
                        {
                            object byrefalias32 = arrReqUnitOptions;
                            _.CALLm1argp(this, _.CALLm0argp(this, objEntry, _.ARGS.Val("PossReqNos")), "Add", _.ARGS.RefIfArray(byrefalias32, _.ARGS.Ref(intIndex, v141 => { intIndex = v141; }), _.ARGS.Val("ReqNo")));
                        }
                        if (!loopConstraintsInitialized)
                            break;
                        var continueLoop = false;
                        _.HANDLEERROR(errOn, () => {
                            intIndex = _.ADD(intIndex, (Int16)1);
                            continueLoop = _.StrictLTE(intIndex, loopEnd16);
                        });
                        if (!continueLoop)
                            break;
                    }
                }
            }

            // If there is at least one possible requirement match, add entry to list
            // (Otherwise, we can't do anything with the selection so don't bother with it)
            if (_.IF(_.GT(_.NullableNUM(_.CALLm1v0(this, _.CALLm0argp(this, objEntry, _.ARGS.Val("PossReqNos")), "Count")), (Int16)0)))
            {
                object byrefalias33 = arrReqUnitSelections;
                _.CALLm1argp(this, byrefalias33, "Add", _.ARGS.Ref(objEntry, v142 => { objEntry = v142; }));
            }

            _.RELEASEERRORTRAPPINGTOKEN(errOn);
            return BookingUI_UnitSel_AddReqUnitSelection_retVal;
        }
        public object BookingUI_UnitSel_GetMatchedReqUnitSelection(ref object arrReqUnitOptions, ref object arrReqUnitSelections)
        {
            object BookingUI_UnitSel_GetMatchedReqUnitSelection_retVal = null;
            object lsPermutations = null;
            object lsTemp = null;
            object lsPossReqNos = null;
            object intIndex = null;
            object intIndexSel = null;
            object intIndexPoss = null;
            object intIndexPerm = null;
            object intIndexOption = null;
            object intScore = null;
            object intBestScore = null;
            object strBestPermutation = null;
            object arrMatches = null;
            object intUnitKey = null;
            object lsUnitKeys = null;
            object GetMatchedReqUnitSelection = null;
            // Given list of requirement option objects and unit selection objects, try to match them up.

            // Ensure we've got values for both lists
            if (_.IF(_.OR(_.ISNULL(arrReqUnitOptions), _.ISNULL(arrReqUnitSelections))))
            {
                GetMatchedReqUnitSelection = VBScriptConstants.Nothing;
            }
            if (_.IF(_.OR(_.EQ(_.NullableNUM(_.CALLm1v0(this, arrReqUnitOptions, "Count")), (Int16)0), _.EQ(_.NullableNUM(_.CALLm1v0(this, arrReqUnitSelections, "Count")), (Int16)0))))
            {
                GetMatchedReqUnitSelection = VBScriptConstants.Nothing;
            }

            // First, create a list of ways in which the unit selections could be applied to the unit
            // options. We'll get out a list of strings which are comma-separated lists; the values
            // will relate the arrReqUnitSelections list indices to arrReqUnitOptions entries.
            //  eg. string "2,3,1"
            //      maps Selection 1 -> Option 2
            //           Selection 2 -> Option 3
            //           Selection 3 -> Option 1
            lsPermutations = _.OBJ(_.CALLm2v1(this, _outer.Page, "Functions", "GetNewObject", "clsList"));
            var loopEnd17 = _.NUM(_.SUBT(_.CALLm1v0(this, arrReqUnitSelections, "Count"), (Int16)1));
            var loopStart17 = _.NUM((Int16)0, loopEnd17, (Int16)1);
            if (_.StrictLTE(loopStart17, loopEnd17))
            {
                for (intIndexSel = loopStart17; _.StrictLTE(intIndexSel, loopEnd17); intIndexSel = _.ADD(intIndexSel, (Int16)1))
                {
                    lsPossReqNos = _.OBJ(_.CALLm0argp(this, _.CALLm0argp(this, arrReqUnitSelections, _.ARGS.Ref(intIndexSel, v143 => { intIndexSel = v143; })), _.ARGS.Val("PossReqNos")));
                    if (_.IF(_.EQ(_.NullableNUM(_.CALLm1v0(this, lsPermutations, "Count")), (Int16)0)))
                    {
                        // This is the first pass, so initialise the permutations list with
                        // the possible matches from this first ReqUnitSelection
                        var loopEnd18 = _.NUM(_.SUBT(_.CALLm1v0(this, lsPossReqNos, "Count"), (Int16)1));
                        var loopStart18 = _.NUM((Int16)0, loopEnd18, (Int16)1);
                        if (_.StrictLTE(loopStart18, loopEnd18))
                        {
                            for (intIndexPoss = loopStart18; _.StrictLTE(intIndexPoss, loopEnd18); intIndexPoss = _.ADD(intIndexPoss, (Int16)1))
                            {
                                _.CALLm1argp(this, lsPermutations, "Add", _.ARGS.RefIfArray(lsPossReqNos, _.ARGS.Ref(intIndexPoss, v144 => { intIndexPoss = v144; })));
                            }
                        }
                    }
                    else
                    {
                        // We want to take our whatever permutation strings we have so far and expand
                        // them to include the possibilities for this ReqUnitSelection
                        // - Make a copy of lsPermutations thus far
                        lsTemp = _.OBJ(_.CALLm2v1(this, _outer.Page, "Functions", "GetNewObject", "clsList"));
                        var loopEnd19 = _.NUM(_.SUBT(_.CALLm1v0(this, lsPermutations, "Count"), (Int16)1));
                        var loopStart19 = _.NUM((Int16)0, loopEnd19, (Int16)1);
                        if (_.StrictLTE(loopStart19, loopEnd19))
                        {
                            for (intIndexPerm = loopStart19; _.StrictLTE(intIndexPerm, loopEnd19); intIndexPerm = _.ADD(intIndexPerm, (Int16)1))
                            {
                                _.CALLm1argp(this, lsTemp, "Add", _.ARGS.RefIfArray(lsPermutations, _.ARGS.Ref(intIndexPerm, v145 => { intIndexPerm = v145; })));
                            }
                        }
                        // - Clear out permutation list
                        _.CALLm1v0(this, lsPermutations, "Clear");
                        // - Re-create new list using previous values with new combinations
                        var loopEnd20 = _.NUM(_.SUBT(_.CALLm1v0(this, lsPossReqNos, "Count"), (Int16)1));
                        var loopStart20 = _.NUM((Int16)0, loopEnd20, (Int16)1);
                        if (_.StrictLTE(loopStart20, loopEnd20))
                        {
                            for (intIndexPoss = loopStart20; _.StrictLTE(intIndexPoss, loopEnd20); intIndexPoss = _.ADD(intIndexPoss, (Int16)1))
                            {
                                var loopEnd21 = _.NUM(_.SUBT(_.CALLm1v0(this, lsTemp, "Count"), (Int16)1));
                                var loopStart21 = _.NUM((Int16)0, loopEnd21, (Int16)1);
                                if (_.StrictLTE(loopStart21, loopEnd21))
                                {
                                    for (intIndexPerm = loopStart21; _.StrictLTE(intIndexPerm, loopEnd21); intIndexPerm = _.ADD(intIndexPerm, (Int16)1))
                                    {
                                        _.CALLm1v1(this, lsPermutations, "Add", _.CONCAT(_.CALLm0argp(this, lsTemp, _.ARGS.Ref(intIndexPerm, v146 => { intIndexPerm = v146; })), ",", _.CALLm0argp(this, lsPossReqNos, _.ARGS.Ref(intIndexPoss, v147 => { intIndexPoss = v147; }))));
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Now determine which arrangement matches the most selection / options pairs
            intBestScore = (Int16)(-1);
            var loopEnd22 = _.NUM(_.SUBT(_.CALLm1v0(this, lsPermutations, "Count"), (Int16)1));
            var loopStart22 = _.NUM((Int16)0, loopEnd22, (Int16)1);
            if (_.StrictLTE(loopStart22, loopEnd22))
            {
                for (intIndex = loopStart22; _.StrictLTE(intIndex, loopEnd22); intIndex = _.ADD(intIndex, (Int16)1))
                {
                    intScore = _.VAL(_.CALLm1argp(this, _outer, "BookingUI_UnitSel_ScoreUnitSelPermutation", _.ARGS.RefIfArray(lsPermutations, _.ARGS.Ref(intIndex, v150 => { intIndex = v150; }))));
                    if (_.IF(_.GT(intScore, intBestScore)))
                    {
                        intBestScore = _.VAL(intScore);
                        strBestPermutation = _.VAL(_.CALLm0argp(this, lsPermutations, _.ARGS.Ref(intIndex, v151 => { intIndex = v151; })));
                    }
                }
            }

            // Finally, translate these matches into UnitKey values (or zero for unit
            // option which don't have a selection matched to them)
            // - Start off with a full-size list (matching size of arrReqUnitOptions) with
            //   with all zero values
            lsUnitKeys = _.OBJ(_.CALLm2v1(this, _outer.Page, "Functions", "GetNewObject", "clsList"));
            var loopEnd23 = _.NUM(_.SUBT(_.CALLm1v0(this, arrReqUnitOptions, "Count"), (Int16)1));
            var loopStart23 = _.NUM((Int16)0, loopEnd23, (Int16)1);
            if (_.StrictLTE(loopStart23, loopEnd23))
            {
                for (intIndex = loopStart23; _.StrictLTE(intIndex, loopEnd23); intIndex = _.ADD(intIndex, (Int16)1))
                {
                    _.CALLm1v1(this, lsUnitKeys, "Add", (Int16)0);
                }
            }

            // - Now push in the selection matches we have
            //    > Split best permutation back into integer values in arrMatches
            //    > The index of arrMatches will matches the index of arrReqUnitSelections
            //    > The value of arrMatches(n) will be the ReqNo it matches, which is the index
            //      of arrReqUnitOptions + 1 (andso also the index of lsUnitKeys + 1 since these
            //      two lists overlay)
            arrMatches = _.SPLIT(strBestPermutation, ",");
            var loopEnd24 = _.UBOUND(arrMatches);
            var loopStart24 = _.NUM((Int16)0, loopEnd24, (Int16)1);
            if (_.StrictLTE(loopStart24, loopEnd24))
            {
                for (intIndexSel = loopStart24; _.StrictLTE(intIndexSel, loopEnd24); intIndexSel = _.ADD(intIndexSel, (Int16)1))
                {
                    intIndexOption = _.SUBT(_.CALLm0argp(this, arrMatches, _.ARGS.Ref(intIndexSel, v152 => { intIndexSel = v152; })), (Int16)1);
                    intUnitKey = _.VAL(_.CALLm0argp(this, _.CALLm0argp(this, arrReqUnitSelections, _.ARGS.Ref(intIndexSel, v153 => { intIndexSel = v153; })), _.ARGS.Val("UnitKey")));
                    _.SETm0a1(this, lsUnitKeys, intIndexOption, _.VAL(intUnitKey));
                }
            }

            // Return matches!
            // There are the same number of values in lsUnitKeys as in arrReqUnitSelections, and
            // each lsUnitKeys(n) is the UnitKey for arrReqUnitSelections(n)
            BookingUI_UnitSel_GetMatchedReqUnitSelection_retVal = _.OBJ(lsUnitKeys);

            return BookingUI_UnitSel_GetMatchedReqUnitSelection_retVal;
        }
        public object BookingUI_UnitSel_ScoreUnitSelPermutation(ref object strPermutation)
        {
            object BookingUI_UnitSel_ScoreUnitSelPermutation_retVal = null;
            object intIndex = null;
            object intScore = null;
            object arrValues = null;
            object lsReqNos = null;
            // Determine a score for the Unit Selection / Option permutations calculated above.
            // Basically, give a score of one for each non-duplicated match.

            lsReqNos = _.OBJ(_.CALLm2v1(this, _outer.Page, "Functions", "GetNewObject", "clsList"));
            arrValues = _.SPLIT(strPermutation, ",");
            intScore = (Int16)0;
            var loopEnd25 = _.UBOUND(arrValues);
            var loopStart25 = _.NUM((Int16)0, loopEnd25, (Int16)1);
            if (_.StrictLTE(loopStart25, loopEnd25))
            {
                for (intIndex = loopStart25; _.StrictLTE(intIndex, loopEnd25); intIndex = _.ADD(intIndex, (Int16)1))
                {
                    if (_.IF(_.NOT(_.CALLm1argp(this, lsReqNos, "Contains", _.ARGS.RefIfArray(arrValues, _.ARGS.Ref(intIndex, v155 => { intIndex = v155; }))))))
                    {
                        intScore = _.ADD(intScore, (Int16)1);
                        _.CALLm1argp(this, lsReqNos, "Add", _.ARGS.RefIfArray(arrValues, _.ARGS.Ref(intIndex, v156 => { intIndex = v156; })));
                    }
                }
            }

            BookingUI_UnitSel_ScoreUnitSelPermutation_retVal = _.VAL(intScore);

            return BookingUI_UnitSel_ScoreUnitSelPermutation_retVal;
        }
        // ====================================================================================================
        // RENDER: Render options for accommodation products (only used with non-precise fuzzy stays)
        // ====================================================================================================
        // SUMMARY: summarise STAYS for this product which match booking criteria
        // [arsAvail]: ADO unit recordset from availability object
        // [adtStartNight]: date of first night of stay
        // [aiReqNumNights]: integer requested num nights
        public object BookingUI_StaySummary(ref object dtReqFirstNight, ref object iReqNights, ref object dtStayFirstNight, ref object iStayNights, ref object pO)
        {
            object BookingUI_StaySummary_retVal = null;

            // Render each stay result with link to further details
            // - 2009-08-10 DWR: Why do we not render this if "_stay" is in the querystring???
            if (_.IF(_.NOTEQ(_.NullableSTR(_.CALLm0argp(this, _outer.Request, _.ARGS.Val("_stay"))), "")))
            {
                return BookingUI_StaySummary_retVal;
            }

            _.CALLm1v1(this, pO, "Write", "<div class=\"StayCandidateList\">");
            _.CALLm1v1(this, pO, "Write", "<div class=\"StayCandidatesTtl\">");
            _.CALLm1v1(this, pO, "Write", _.CONCAT("<p>", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/flexiblesearchresults", "Flexible Search Results"), "</p>"));
            _.CALLm1v1(this, pO, "Write", "</div>");
            if (_.IF(_.OR(_.NOTEQ(dtStayFirstNight, dtReqFirstNight), _.NOTEQ(iReqNights, iStayNights))))
            {
                _.CALLm1v1(this, pO, "Write", "<div class=\"cell\">");
                _.CALLm1v1(this, pO, "Write", "<div class=\"pnStayTtl\">");
                object byrefalias34 = dtStayFirstNight, byrefalias35 = iStayNights;
                try
                {
                    _.CALLm1v1(this, pO, "Write", _.CALLm1argp(this, _outer, "BookingUI_StayTtl", _.ARGS.Ref(byrefalias34, v157 => { byrefalias34 = v157; }).Ref(byrefalias35, v158 => { byrefalias35 = v158; })));
                }
                finally { dtStayFirstNight = byrefalias34; iStayNights = byrefalias35; }
                _.CALLm1v1(this, pO, "Write", "</div>");
                object byrefalias36 = dtReqFirstNight, byrefalias37 = dtStayFirstNight, byrefalias38 = iReqNights, byrefalias39 = iStayNights;
                try
                {
                    _.CALLm1v1(this, pO, "Write", _.CALLm1argp(this, _outer, "BookingUI_StayDiff", _.ARGS.Ref(byrefalias36, v159 => { byrefalias36 = v159; }).Ref(byrefalias37, v160 => { byrefalias37 = v160; }).Ref(byrefalias38, v161 => { byrefalias38 = v161; }).Ref(byrefalias39, v162 => { byrefalias39 = v162; })));
                }
                finally { dtReqFirstNight = byrefalias36; dtStayFirstNight = byrefalias37; iReqNights = byrefalias38; iStayNights = byrefalias39; }
                _.CALLm1v1(this, pO, "Write", "</div>");
            }
            _.CALLm1v1(this, pO, "Write", "</div>");

            return BookingUI_StaySummary_retVal;
        }
        // SUMMARY: Render details for a single stay, including UNIT booking UI
        // [objAvailEntry]: A single supplier's availability data for a single stay (AvailabilityStayResultsWrapped)
        // [iStayNum]: Only applies when displaying multiple fuzzy results
        // [adtStartNight]: date of first night of stay
        // [aiReqNights]: integer requested num nights
        // [bTeleBooking]: does the current product only support telephone booking (ie. is on tele booking channel)?
        // [strProductBookingWebIfAny]: the Booking Website for the the current product, if there is one (so may be empty, null, blank, whatever)
        // [strEviivoIdIfAny]: the Eviivo Id for the the current product, if there is one (so may be empty, null, blank, whatever)
        // [intProductKey]
        // [strChannel]
        // [bIndicative]: does the specified stay have any indicative units?
        // [bIndicativeValid]: are we within the timeout period for indicative bookings?
        // [lsRemoteUnitSelections]: data regarding unit pre-selections (see VB Deep Linking)
        public object BookingUI_StayDetails(object objAvailEntry, object iStayNum, object adtStartNight, object aiReqNights, object bTeleBooking, object strProductBookingWebIfAny, object strEviivoIdIfAny, object intProductKey, object strChannel, object bIndicative, object bIndicativeValid, object objAllUnits, object lsRemoteUnitSelections, object pO, object bRenderMaximumUnitsAvailable)
        {
            object BookingUI_StayDetails_retVal = null;
            object intIndexUnit = null;
            object objUnit = null;
            object iLastReqmnt = null;
            object iThisReqmnt = null;
            object bGotOpenReqContainer = null;
            object sClassName = null;
            object bPrecise = null;
            object iUnitKey = null;
            object iMaxRq = null;
            object iRemoteUnitKey = null;
            object bSelected = null;
            object strNonBookableUnits = null;
            object bHasBookableUnits = null;
            object bHasNonBookableUnits = null;

            // Ensure we've actually got some availability (we should if we've got here!)
            if (_.IF(_.EQ(_.NullableNUM(_.CALLm2v0(this, objAvailEntry, "Units", "Count")), (Int16)0)))
            {
                _.CALLm1v1(this, _outer.Page, "PrintTraceWarning", "BookingUI_StayDetails: No units in objAvailEntry");
                return BookingUI_StayDetails_retVal;
            }

            // This method opens a new div - we'll need to close it later
            _.CALLm1argp(this, _outer, "BookingUI_RenderNewStay", _.ARGS.Ref(objAvailEntry, v163 => { objAvailEntry = v163; }).Ref(iStayNum, v164 => { iStayNum = v164; }).Ref(adtStartNight, v165 => { adtStartNight = v165; }).Ref(aiReqNights, v166 => { aiReqNights = v166; }).Ref(pO, v167 => { pO = v167; }));

            iMaxRq = (Int16)0;
            iLastReqmnt = (Int16)0;
            iRemoteUnitKey = (Int16)0;
            bGotOpenReqContainer = false;
            bHasBookableUnits = false;
            bHasNonBookableUnits = false;
            var loopEnd26 = _.NUM(_.SUBT(_.CALLm2v0(this, objAvailEntry, "Units", "Count"), (Int16)1));
            var loopStart26 = _.NUM((Int16)0, loopEnd26, (Int16)1);
            if (_.StrictLTE(loopStart26, loopEnd26))
            {
                for (intIndexUnit = loopStart26; _.StrictLTE(intIndexUnit, loopEnd26); intIndexUnit = _.ADD(intIndexUnit, (Int16)1))
                {
                    objUnit = _.OBJ(_.CALLm2argp(this, objAvailEntry, "Units", "GetItem", _.ARGS.Ref(intIndexUnit, v168 => { intIndexUnit = v168; })));

                    iThisReqmnt = _.VAL(_.CALLm1v0(this, objUnit, "ReqNo"));
                    if (_.IF(_.GT(iThisReqmnt, iMaxRq)))
                    {
                        // Moved on to next requirement, get key of pre-selected unit - iRemoteUnitKey
                        // will be zero if no selection has been passed in (applies to deep-linking)
                        iMaxRq = _.VAL(iThisReqmnt);
                        iRemoteUnitKey = _.VAL(_.CALLm1argp(this, _outer, "BookingUI_GetPreSelectedUnitKey", _.ARGS.Ref(lsRemoteUnitSelections, v169 => { lsRemoteUnitSelections = v169; }).Ref(iThisReqmnt, v170 => { iThisReqmnt = v170; })));
                    }

                    // Check whether we're moving into a new requirement (if so, default to having
                    // the first unit appear selected) and render the "Room 1 - for 1 Guest"
                    // content
                    if (_.IF(_.NOTEQ(iThisReqmnt, iLastReqmnt)))
                    {

                        // If we've already got one of these containers open, close its tags
                        if (_.IF(bGotOpenReqContainer))
                        {
                            _.CALLm1v1(this, pO, "Write", "</div></div>");
                        }
                        _.CALLm1argp(this, _outer, "BookingUI_RenderNewReq", _.ARGS.Ref(objUnit, v171 => { objUnit = v171; }).Ref(iStayNum, v172 => { iStayNum = v172; }).Ref(iThisReqmnt, v173 => { iThisReqmnt = v173; }).Val(_.NOT(_.CALLm1v0(this, objAvailEntry, "IsLocal"))).Ref(pO, v174 => { pO = v174; }));
                        bGotOpenReqContainer = true;

                        bSelected = true;
                        iLastReqmnt = _.VAL(iThisReqmnt);
                    }
                    else
                    {
                        bSelected = false;
                    }

                    // .. however, if there was a pre-selected unit key passed in, this should override which
                    // unit appears selected (this only applies when iRemoteUnitKey is not zero, meaning that
                    // a unit selection exists - note: eviivo units always appear with unit key zero)
                    iUnitKey = _.VAL(_.CALLm1v0(this, objUnit, "UnitKey"));
                    if (_.IF(_.NOTEQ(_.NullableNUM(iRemoteUnitKey), (Int16)0)))
                    {
                        bSelected = _.VAL(_.EQ(iUnitKey, iRemoteUnitKey));
                    }

                    // build up a list of invalid indicative or telephone booking
                    // units, this is used later by javascript when we have a mixture of allocated and indicative
                    // availability
                    if (_.IF(_.OR(_.AND(_.CALLm1v0(this, objUnit, "Indicative"), _.NOT(bIndicativeValid)), bTeleBooking)))
                    {
                        bHasNonBookableUnits = true;
                        if (_.IF(_.GT(_.NullableNUM(_.LEN(strNonBookableUnits)), (Int16)0)))
                        {
                            strNonBookableUnits = _.CONCAT(strNonBookableUnits, ",");
                        }

                        //MJ - 	the stay num is no longer part of this data, it is part of each array's name
                        //		look at TB's other changes to see the reasoning behind this
                        strNonBookableUnits = _.CONCAT(strNonBookableUnits, iUnitKey);
                        _.CALLm1v1(this, _outer.Page, "PrintTrace", _.CONCAT("strNonBookableUnits", strNonBookableUnits));
                    }
                    else
                    {
                        bHasBookableUnits = true;
                    }

                    // 2009-09-30 DWR: The AvailClassName was previously generated by considering the indicative
                    // state of the whole stay - this was causing all units to be rendered as indicative if any
                    // one of them was, now we take the indicative state from each unit (but keep the indicative
                    // "validity" from the whole stay, where required)
                    _.CALLm1argp(this, _outer, "BookingUI_RenderUnit", _.ARGS.Ref(iStayNum, v175 => { iStayNum = v175; }).Ref(iThisReqmnt, v176 => { iThisReqmnt = v176; }).Ref(bSelected, v177 => { bSelected = v177; }).Ref(objAvailEntry, v178 => { objAvailEntry = v178; }).Ref(objUnit, v179 => { objUnit = v179; }).Ref(objAllUnits, v180 => { objAllUnits = v180; }).Val(_.CALLm1argp(this, _outer, "BookingUI_AvailClassName", _.ARGS.Val(_.CALLm1v0(this, objUnit, "Indicative")).Ref(bIndicativeValid, v181 => { bIndicativeValid = v181; }).Ref(bTeleBooking, v182 => { bTeleBooking = v182; }))).Ref(pO, v183 => { pO = v183; }).Ref(bRenderMaximumUnitsAvailable, v184 => { bRenderMaximumUnitsAvailable = v184; }));

                }
            }

            // Ensure any open req container (eg. "Room 1 - for 1 Guest" section) is closed
            if (_.IF(bGotOpenReqContainer))
            {
                _.CALLm1v1(this, pO, "Write", "</div></div>");
                bGotOpenReqContainer = false;
            }

            // Close the BookingUI_RenderNewStay containing div
            _.CALLm1v1(this, pO, "Write", "</div>");

            // Wrap these hidden inputs in a div for html validity
            _.CALLm1v1(this, pO, "Write", "<div>");
            _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"_nStays\" value=\"", iStayNum, "\" />"));
            _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"_nReqs\" value=\"", iMaxRq, "\" />"));
            if (_.IF(_.NOT(_.CALLm1v0(this, objAvailEntry, "IsLocal"))))
            {
                _.CALLm1v1(this, pO, "Write", "<input type=\"hidden\" name=\"IsEviivoBooking\" value=\"yes\" />");
                if (_.IF(_outer.IsExternalBooking))
                {
                    _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"eviivoconf\" value=\"", _.CLNG(_.CONCAT("0", _.CALLm2v1(this, _outer.Page, "Site", "Params", "Integration_Eviivo_ConfigSet"))), "\" />"));
                }
            }
            _.CALLm1v1(this, pO, "Write", "</div>");

            // 2014-06-25 DWR: For sites that use the legacy "eviivo external" booking integration (meaning sites where VB Polling is not enabled - the new implementation
            // results in Eviivo results being reported as Polling results and the user being sent through the Polling Exit with a fully-populated deep link), the Book
            // button should not be shown here. The Unit Selection should never be shown in this case, to be honest, since Book buttons should go straight to the Product's
            // Booking Website and not enter the site's availability process. However, if there are sites that show inline Unit Selection (inline with the Product List)
            // then the Unit data may be useful. If we were wanted to render Book buttons here (to the external site) then logic would have to be duplicated from the
            // Product List or Detail Control, which would be better avoided. A much better solution is to enable VB Polling and avoid this legacy mechanism entirely.
            // Note: We could potentially render the button for Local Avail and not for Eviivo but I think that that's more confusing than helpful, particularly since
            // it's inconsistent with the Product List / Detail implementation (which bases its decision upon whether the Product has an Eviivo Id).
            if (_.IF(_.AND(_.AND(_.NOT(_.CALLm2v1(this, _outer.Page, "Site", "Params", "Booking_EnablePolling")), _.NOTEQ(_.NullableSTR(_.TRIM(_.CONCAT("", strEviivoIdIfAny))), "")), _.CALLm2v1(this, _outer.Page, "Site", "Params", "Integration_Eviivo_ExtBooking_Enable"))))
            {
                _.CALLm1v1(this, _outer.Page, "PrintTraceWarning", "Not rendering any Book buttons for Unit Selection since the legacy Eviivo External Booking configuration is enabled (the recommended alternative is to use the deep-link-supporting Eviivo External Booking configuration, this may be done by enabling VB Polling)");
                return BookingUI_StayDetails_retVal;
            }

            // 2014-03-14 DWR: New functionality "Availability Searches with offsite Booking Web Booking" allows for Products to be on the Telephone Booking Channel
            // and have their availability queried but to show a Booking button that goes to the Product's Booking Website (if one is specified), rather than
            // showing a "this can not be booked online, please call.." message (this means that the avail criteria have to be re-entered on the target
            // website, but that is understood and how it works - see FogBugz 10367). I've tried to make the markup for this button reminiscent of
            // that in Product List and Detail to try to make any additional styling requirements as low as possible.
            strProductBookingWebIfAny = _.VAL(_.TRIM(_.CONCAT("", strProductBookingWebIfAny)));
            if (_.IF(_.AND(_.AND(_.AND(bTeleBooking, _.CALLm2v1(this, _outer.Page, "Site", "Params", "Booking_EnableByPhone")), _.CALLm2v1(this, _outer.Page, "Site", "Params", "Booking_AllowOffSiteTelephoneBookings")), _.NOTEQ(_.NullableSTR(strProductBookingWebIfAny), ""))))
            {
                _.CALLm1v1(this, _outer.Page, "PrintTrace", "Since this is a Telephone Booking Product with a Booking Website and the 'Allow Offsite Booking Web Booking for Telephone Bookings' parameter is enabled, a button to the Booking Website is being rendered");
                _.CALLm1v1(this, pO, "Write", "<div class=\"pnStayButtons\">");
                _.CALLm1v1(this, pO, "Write", "<p class=\"bookonline\">");
                _.CALLm1v1(this, pO, "Write", "<a href=\"");
                _.CALLm1v1(this, pO, "Write", _.CALLm1argp(this, _outer.Server, "HtmlEncode", _.ARGS.Ref(strProductBookingWebIfAny, v185 => { strProductBookingWebIfAny = v185; })));
                _.CALLm1v1(this, pO, "Write", "\"");
                if (_.IF(_.OR(_.CALLm1v0(this, _outer.Page, "IsPartialRender"), _.EQ(_.NullableSTR(_.CALLm0argp(this, _outer.Request, _.ARGS.Val("PartialRenderType"))), "html"))))
                {
                    // If in Partial Render then set target="_blank" instead of rel="external" (we only do the latter for strict adherence to standards and then
                    // use javascript to transform after rendering - when requesting additional content through javascript this transformation won't be performed
                    // so we'll need to generate it direct)
                    // 2014-06-12 DWR: The partial render requests for this data are commonly made as "html" meaning that Page.IsPartialRender will be false
                    // (the logic being that Controls should render entirely as standard when in html partial render mode) so I've added an additional check
                    // for the a "PartialRenderType" value of "html" to ensure that the new-window logic is maintained correctly.
                    _.CALLm1v1(this, pO, "Write", " target=\"_blank\"");
                }
                else
                {
                    _.CALLm1v1(this, pO, "Write", " rel=\"external\"");
                }
                _.CALLm1v1(this, pO, "Write", " class=\"ProvClickCustom\" name=\"PROBWEBREF|");
                // This is the "Provider Booking Website Referral" statistic, as required by the SharePoint document for FogBugz 10367
                _.CALLm1v1(this, pO, "Write", _.CALLm1argp(this, _outer.Server, "HtmlEncode", _.ARGS.Ref(strChannel, v186 => { strChannel = v186; })));
                _.CALLm1v1(this, pO, "Write", "|");
                _.CALLm1argp(this, pO, "Write", _.ARGS.Ref(intProductKey, v187 => { intProductKey = v187; }));
                _.CALLm1v1(this, pO, "Write", "\"");
                _.CALLm1v1(this, pO, "Write", ">");
                _.CALLm1v1(this, pO, "Write", "<img src=\"");
                _.CALLm1v1(this, pO, "Write", _.CALLm1v2(this, _outer.Page, "ImageResource", "bookonline/btn/book", _.CONCAT(_.CALLm1v0(this, _outer.Context, "ImageDir"), "booking/book.gif")));
                _.CALLm1v1(this, pO, "Write", "\" alt=\"");
                _.CALLm1v1(this, pO, "Write", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/btn/book", "Book"));
                _.CALLm1v1(this, pO, "Write", " (");
                _.CALLm1v1(this, pO, "Write", _.CALLm1v2(this, _outer.Page, "Resource", "productdetail/bookonline/opensinanewwindow", "opens in a new window"));
                _.CALLm1v1(this, pO, "Write", ")\" ");
                _.CALLm1v1(this, pO, "Write", "/>");
                _.CALLm1v1(this, pO, "Write", "</a>");
                _.CALLm1v1(this, pO, "Write", "</p>");
                _.CALLm1v1(this, pO, "Write", _.CONCAT("</div>", VBScriptConstants.vbCrLf));
                return BookingUI_StayDetails_retVal;
            }

            // 2014-03-13 DWR: If there is at least one bookable unit then display the Book button and rely on JavaScript to show/hide it if selections are made that
            // can not be completed online. But if there are NO bookable units (eg. a Telephone Booking Product or all of the Units are Indicative where the timeout
            // period has passed) then there's no point even rendering the button.
            if (_.IF(bHasBookableUnits))
            {
                _.CALLm1argp(this, _outer, "BookingUI_RenderButtons", _.ARGS.Ref(iStayNum, v188 => { iStayNum = v188; }).Ref(pO, v189 => { pO = v189; }).Val(_.CALLm1v0(this, objAvailEntry, "IsExternal")));
            }

            // if we have an invalid indicative unit or telephone unit then
            // render this message - let the js do the rest
            if (_.IF(bHasNonBookableUnits))
            {

                // 2010-07-09 PW: RIP Gary
                // This is the array formerly known as garyTeleBookUnitKeys
                // it is used for switching between the online book button if the unit is bookable
                // or rendering the relevant warning message if it isn't
                // 2010-10-21 TB: augmenting Gary with stay key. This is to allow for multiple stays
                // in which this JS is executed on a per stay basis via a partial render.
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<script type=\"text/javascript\">", VBScriptConstants.vbCrLf));
                _.CALLm1v1(this, pO, "Write", _.CONCAT(" var aryNonBookableUnits_", iStayNum, " = [", strNonBookableUnits, "]; ", VBScriptConstants.vbCrLf));
                _.CALLm1v1(this, pO, "Write", _.CONCAT(" var iTotalNonBookableUnits = ", iThisReqmnt, ";", VBScriptConstants.vbCrLf));
                _.CALLm1v1(this, pO, "Write", _.CONCAT("</", "script>"));

                // Render relevant offline booking message
                _.CALLm1v1(this, pO, "Write", "<div id=\"pnTeleBook_PromptCall\">");
                if (_.IF(_.CALLm2v1(this, _outer.Page, "Site", "Params", "Booking_EnableByPhone")))
                {
                    _.CALLm1v1(this, pO, "Write", _.CONCAT("<p>", _.REPLACE(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/telebook/prompt", "One or more of the units you have selected must be booked via telephone. Please ring #bookingtelephone# to continue this booking."), "#bookingtelephone#", _.CALLm2v1(this, _outer.Page, "Site", "Params", "Booking_TelephoneNumber")), "</p>"));
                }
                else
                {
                    _.CALLm1v1(this, pO, "Write", _.CONCAT("<p>", _.REPLACE(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/indtelebook/prompt", "Although available, some of the units you have selected cannot be booked online. Alternatively, select different units with online booking only."), "#bookingtelephone#", _.CALLm2v1(this, _outer.Page, "Site", "Params", "Booking_TelephoneNumber")), "</p>"));
                }
                _.CALLm1v1(this, pO, "Write", "</div>");
            }

            return BookingUI_StayDetails_retVal;
        }
        public object BookingUI_GetPreSelectedUnitKey(object lsRemoteUnitSelections, object iReqNo)
        {
            object BookingUI_GetPreSelectedUnitKey_retVal = null;
            // remote referrals (eg. VB integration) will include a UNIT_KEY choice and CHILD_COUNT.
            // eg. Request vars formatted as 'URslt[REQ_NUMBER]=[UNIT_KEY]-[NUM_ADULT]-[NUM-CHILD]'

            // If not got here from a remote referral (eg. VB integration), the lsRemoveUnitSelections will be Nothing
            if (_.IF(_.NOT(_.IS(lsRemoteUnitSelections, VBScriptConstants.Nothing))))
            {
                // Get unit selection passed in (may be zero if invalid request was made)
                // Note: lsRemoteUnitSelections has zero-based index, iReqNo is one-based
                if (_.IF(_.AND(_.GTE(_.NullableNUM(iReqNo), (Int16)1), _.LTE(iReqNo, _.CALLm1v0(this, lsRemoteUnitSelections, "Count")))))
                {
                    BookingUI_GetPreSelectedUnitKey_retVal = _.VAL(_.CALLm0argp(this, lsRemoteUnitSelections, _.ARGS.Val(_.SUBT(iReqNo, (Int16)1))));
                    return BookingUI_GetPreSelectedUnitKey_retVal;
                }
            }

            BookingUI_GetPreSelectedUnitKey_retVal = (Int16)0;
            return BookingUI_GetPreSelectedUnitKey_retVal;
        }
        // SUMMARY: for VB Polling - we want to render a supplier name and icon above each set of unit options
        public object BookingUI_StayDetails_PollingHeader(object objAvailEntry, object pO, object strSupplierLogo, object strSupplierName)
        {
            object BookingUI_StayDetails_PollingHeader_retVal = null;

            // Render header content (icon, if specified) and supplier name
            // 2008-12-18 DWR: Add a style to indicate whether supplier is Local, FrontDesk or External (this will
            // allow a custom logo to be used for Local or FrontDesk, for example)
            _.CALLm1v1(this, pO, "Write", "<div class=\"StayCandidateItemHeader ");
            if (_.IF(_.CALLm1v0(this, objAvailEntry, "IsLocal")))
            {
                _.CALLm1v1(this, pO, "Write", " AvailLocal");
            }
            else if (_.IF(_.CALLm1v0(this, objAvailEntry, "IsRemote")))
            {
                _.CALLm1v1(this, pO, "Write", " AvailFrontDesk");
            }
            else
            {
                _.CALLm1v1(this, pO, "Write", " AvailExternal");
            }
            _.CALLm1v1(this, pO, "Write", "\">");
            if (_.IF(_.NOTEQ(_.NullableSTR(strSupplierLogo), "")))
            {
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<img src=\"", strSupplierLogo, "\" alt=\"", strSupplierName, "\" />"));
            }
            _.CALLm1v1(this, pO, "Write", _.CONCAT("<h2>", strSupplierName, "</h2>"));
            _.CALLm1v1(this, pO, "Write", _.CONCAT("</div>", VBScriptConstants.vbCrLf));
            return BookingUI_StayDetails_PollingHeader_retVal;
        }
        //tries to get a supplier logo for us
        public object GetSupplierLogo(ref object strProductEstateID)
        {
            object GetSupplierLogo_retVal = null;
            object strSupplierLogo = null;
            strSupplierLogo = "";
            if (_.IF(_outer.IsExternalBooking))
            {
                strSupplierLogo = _.VAL(_.CALLm1v2(this, _outer.Page, "ImageResource", _.CONCAT("bookonline/unitselection/polling/localsupplier/estate_", strProductEstateID, "/logo"), ""));
                if (_.IF(_.EQ(_.NullableSTR(strSupplierLogo), "")))
                {
                    strSupplierLogo = _.VAL(_.CALLm1v2(this, _outer.Page, "Resource", _.CONCAT("bookonline/unitselection/polling/localsupplier/estate_", strProductEstateID, "/logo"), ""));
                    if (_.IF(_.NOTEQ(_.NullableSTR(strSupplierLogo), "")))
                    {
                        _.CALLm1v1(this, _outer.Page, "PrintTraceWarning", "Loaded estate scoped supplier logo from a deprecated location - please move it to the image resources language file");
                    }
                }
            }
            if (_.IF(_.EQ(_.NullableSTR(strSupplierLogo), "")))
            {
                strSupplierLogo = _.VAL(_.CALLm1v2(this, _outer.Page, "ImageResource", "bookonline/unitselection/polling/localsupplier/logo", ""));
                if (_.IF(_.EQ(_.NullableSTR(strSupplierLogo), "")))
                {
                    strSupplierLogo = _.VAL(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/polling/localsupplier/logo", ""));
                    if (_.IF(_.NOTEQ(_.NullableSTR(strSupplierLogo), "")))
                    {
                        _.CALLm1v1(this, _outer.Page, "PrintTraceWarning", "Loaded estate scoped supplier logo from a deprecated location - please move it to the image resources language file");
                    }
                }
            }
            GetSupplierLogo_retVal = _.VAL(strSupplierLogo);
            return GetSupplierLogo_retVal;
        }
        // SUMMARY: return URL which browsers without Javascript can use to navigate stay candidates page
        // [aiStay]: integer stay number. 1 = 1st stay, 2 = 2nd stay. Zero produces back URL to stay candidates page
        // <retval>: string URL for hyperlink
        public object BookingUI_StayDetailsUrl(ref object aiStay)
        {
            object BookingUI_StayDetailsUrl_retVal = null;
            object sUrl = null;
            object sStay = null;
            object iPos = null;
            object sRight = null;
            object iRight = null; /* Undeclared in source */

            // get current URL. Prepare [stay] variable to be appended to URL
            sUrl = _.VAL(_.CALLm1v1(this, _outer.Request, "ServerVariables", "HTTP_X_REWRITE_URL"));
            if (_.IF(_.GT(_.NullableNUM(aiStay), (Int16)0)))
            {
                sStay = _.CONCAT("&_stay=", aiStay);
            }
            else
            {
                sStay = "";
            }

            // does URL already have stay variable? if so, remove it and return new URL
            iPos = _.VAL(_.INSTR(sUrl, "&_stay="));
            if (_.IF(_.GT(_.NullableNUM(iPos), (Int16)0)))
            {
                sRight = _.VAL(_.MID(sUrl, _.ADD(iPos, (Int16)7)));
                iRight = _.VAL(_.INSTR(sRight, "&"));
                sUrl = _.VAL(_.LEFT(sUrl, _.SUBT(iPos, (Int16)1)));
                if (_.IF(_.GT(_.NullableNUM(iRight), (Int16)0)))
                {
                    sUrl = _.CONCAT(sUrl, _.MID(sRight, iRight));
                }
            }
            BookingUI_StayDetailsUrl_retVal = _.CONCAT(sUrl, sStay);
            return BookingUI_StayDetailsUrl_retVal;
        }
        // SUMMARY: render new stay UI - WARNING: this doesn't close all of the elements it opens!
        // [objAvailEntry]: avail data for a single stay
        // [aiStayNum]: integer stay index (1-based)
        // [adtStartNight]: date requested start night
        // [aiReqNights]: integer requested num nights
        public object BookingUI_RenderNewStay(object objAvailEntry, object aiStayNum, object adtStartNight, object aiReqNights, object pO)
        {
            object BookingUI_RenderNewStay_retVal = null;
            object sPostfix = null;
            object bPrecise = null;
            object bExactMatch = null;

            // Render slightly differently if got a precise match
            // - Also render differently when VB Polling enabled, since we have to render
            //   more of these sections than otherwise
            bExactMatch = _.VAL(_.AND(_.EQ(_.CALLm1v0(this, objAvailEntry, "StartDate"), adtStartNight), _.EQ(_.CALLm1v0(this, objAvailEntry, "Nights"), aiReqNights)));
            if (_.IF(_.OR(bExactMatch, _outer.IsVBPollingEnabled)))
            {
                bPrecise = true;
                sPostfix = "1";
            }
            else if (_.IF(_.EQ(_.CLNG(_.CONCAT("0", _.CALLm0argp(this, _outer.Request, _.ARGS.Val("_stay")))), aiStayNum)))
            {
                sPostfix = "1";
            }
            else
            {
                sPostfix = "";
            }

            // If not exact match then render a warning as well as the date difference later
            if (_.IF(_.NOT(bExactMatch)))
            {
                _.CALLm1argp(this, _outer, "RenderNotRequiredDateWarning", _.ARGS.Ref(pO, v190 => { pO = v190; }));
            }

            _.CALLm1v1(this, pO, "Write", _.CONCAT("<div class=\"StayCandidateItem", sPostfix, "\">", VBScriptConstants.vbCrLf));

            if (_.IF(_.NOT(bExactMatch)))
            {
                _.CALLm1v1(this, pO, "Write", "<div class=\"pnStayTtl\">");
                _.CALLm1v1(this, pO, "Write", "<p>");
                _.CALLm1v1(this, pO, "Write", _.CALLm1v2(this, _outer, "BookingUI_StayTtl", _.CALLm1v0(this, objAvailEntry, "StartDate"), _.CALLm1v0(this, objAvailEntry, "Nights")));
                _.CALLm1v1(this, pO, "Write", "</p>");
                _.CALLm1v1(this, pO, "Write", _.CONCAT("</div>", VBScriptConstants.vbCrLf));
                if (_.IF(_.NOT(_outer.bRenderAsCalendar)))
                {
                    _.CALLm1v1(this, pO, "Write", _.CALLm1argp(this, _outer, "BookingUI_StayDiff", _.ARGS.Ref(adtStartNight, v191 => { adtStartNight = v191; }).Val(_.CALLm1v0(this, objAvailEntry, "StartDate")).Ref(aiReqNights, v192 => { aiReqNights = v192; }).Val(_.CALLm1v0(this, objAvailEntry, "Nights"))));
                }
            }
            return BookingUI_RenderNewStay_retVal;
        }
        // SUMMARY: return title for this stay candidate
        // [aiNights]: integer number nights for this stay
        // [adtFirstNight]: date of first night
        // [adtLastNight]: date of last night
        // <retval>: string stay title
        public object BookingUI_StayTtl(object adtFirstNight, object aiNights)
        {
            object BookingUI_StayTtl_retVal = null;
            if (_.IF(_.EQ(_.NullableNUM(aiNights), (Int16)1)))
            {
                BookingUI_StayTtl_retVal = _.CONCAT(aiNights, _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/nightstart", " night, start "), _.CALLm3argp(this, _outer.Page, "Functions", "Dates", "ShortDate", _.ARGS.Ref(adtFirstNight, v193 => { adtFirstNight = v193; })));
                return BookingUI_StayTtl_retVal;
            }

            BookingUI_StayTtl_retVal = _.CONCAT(aiNights, _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/nightsfrom", " nights, from "), _.CALLm3argp(this, _outer.Page, "Functions", "Dates", "ShortDate", _.ARGS.Ref(adtFirstNight, v195 => { adtFirstNight = v195; })), _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/to", " to "), _.CALLm3v1(this, _outer.Page, "Functions", "Dates", "Shortdate", _.DATEADD("d", aiNights, adtFirstNight)));
            return BookingUI_StayTtl_retVal;
        }
        // SUMMARY: describe difference between THIS DATE and REQUESTED stay date
        // [adtReqDate]: date of REQUESTED first night of stay
        // [adtThisDate]: date of RESULTANT first night of stay
        // [aiReqNights]: integer requested num nights
        // [aiNights]: integer result num nights
        public object BookingUI_StayDiff(object adtReqDate, object adtThisDate, object aiReqNights, object aiResultNights)
        {
            object BookingUI_StayDiff_retVal = null;
            object iDateDiff = null;
            object iDurDiff = null;

            iDateDiff = _.VAL(_.DATEDIFF("d", adtReqDate, adtThisDate));
            iDurDiff = _.SUBT(aiResultNights, aiReqNights);
            BookingUI_StayDiff_retVal = _.CONCAT("<div class=\"pnStayDiff\">", _.CALLm3argp(this, _outer.Page, "Functions", "Booking", "Booking_MatchQual", _.ARGS.Val((Int16)0).Ref(iDateDiff, v197 => { iDateDiff = v197; }).Ref(iDurDiff, v198 => { iDurDiff = v198; }).Ref(aiReqNights, v199 => { aiReqNights = v199; }).Val((Int16)2)), "</div>", VBScriptConstants.vbCrLf);
            return BookingUI_StayDiff_retVal;
        }
        // SUMMARY: render new requirement UI - WARNING: this doesn't close all of the elements it opens!
        // [arsAvail]: ADO unit recordset from availability object
        // [aiStayNum]: integer stay index
        // [aiThisReqmnt]: integer requirement number (from recordset)
        public object BookingUI_RenderNewReq(object objUnit, object aiStayNum, object aiThisReqmnt, object abRemote, object pO)
        {
            object BookingUI_RenderNewReq_retVal = null;
            object iSz = null;
            object sReqmntSet = null;
            object sRemoteRqmnt = null;
            object iChild = null;
            object iRemoteNumChild = null;
            object sSelected = null;
            object aryChildAges = null;
            object iChildAgeIndex = null;

            iSz = _.VAL(_.CALLm1v0(this, objUnit, "ReqSize"));

            _.CALLm1v1(this, pO, "Write", _.CONCAT("<div class=\"pnStayReqmnt\">", VBScriptConstants.vbCrLf));
            _.CALLm1v1(this, pO, "Write", "<div class=\"pnStayReqmntTtl\">");
            _.CALLm1v1(this, pO, "Write", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/room", "Room"));
            _.CALLm1v1(this, pO, "Write", " ");
            _.CALLm1argp(this, pO, "Write", _.ARGS.Ref(aiThisReqmnt, v203 => { aiThisReqmnt = v203; }));
            _.CALLm1v1(this, pO, "Write", " - ");
            _.CALLm1v1(this, pO, "Write", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/for", "for"));
            _.CALLm1v1(this, pO, "Write", " ");
            _.CALLm1argp(this, pO, "Write", _.ARGS.Ref(iSz, v204 => { iSz = v204; }));
            _.CALLm1v1(this, pO, "Write", " ");
            _.CALLm1v1(this, pO, "Write", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/guest(s)", "guest(s)"));

            //#MJ -	We can only render our room requirement data based upon the recieved dat, not the requirement we passed in, as it may have been fulfilled in a different order
            //2012-03-29 NP: Here we render the requirements that are linked to the unit stay details in the response from the Avail Component
            // we do NOT want to render the original request against each unit that is rendered because they may not order up
            // Example: Request roomReq_1 = 2; roomReq_2 = 1; Response may come back in a different order
            // i.e. unit_1 with ReqSize = 1, unit_2 with ReqSize = 2 so roomReq_1 = 1, roomReq= 2; they end up swapped around
            _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"roomReq_", aiThisReqmnt, "\" value=\"", iSz, "\" />"));

            //#MJ - need to check with Rich if we want to indicate who's going into what room
            if (_.IF(_.AND(_.CALLm2v1(this, _outer.Page, "Site", "Params", "Booking_ChildPricing"), _.GT(_.NullableNUM(_.CALLm1v0(this, objUnit, "ChildrenRequirement")), (Int16)0))))
            {
                _.CALLm1v1(this, pO, "Write", " - (");
                _.CALLm1v1(this, pO, "Write", "<span class=\"ReqmntDetails\">");
                _.CALLm1v1(this, pO, "Write", _.CALLm1v2(this, _outer.Page, "Resource", "adults", "Adults"));
                _.CALLm1v1(this, pO, "Write", ": ");
                _.CALLm1v1(this, pO, "Write", _.CALLm1v0(this, objUnit, "AdultsRequirement"));
                _.CALLm1v1(this, pO, "Write", " ");
                _.CALLm1v1(this, pO, "Write", _.CALLm1v2(this, _outer.Page, "Resource", "children", "Children"));
                _.CALLm1v1(this, pO, "Write", ": ");
                _.CALLm1v1(this, pO, "Write", _.CALLm1v0(this, objUnit, "ChildrenRequirement"));
                _.CALLm1v1(this, pO, "Write", ") ");
                _.CALLm1v1(this, pO, "Write", "</span>");
                // NP 2012-03-01: Child pricing requirements were not previously being posted to the checkout
                // Adult & Child Requirement amount is needed by the RequirementSummary control and the child ages are
                // needed by the checkout for creating the correct requirement record with the relevant discount values
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"roomReq_", aiThisReqmnt, "_adults\" value=\"", _.CALLm1v0(this, objUnit, "AdultsRequirement"), "\" />"));
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"roomReq_", aiThisReqmnt, "_children\" value=\"", _.CALLm1v0(this, objUnit, "ChildrenRequirement"), "\" />"));

                // ChildrenAges is a comma separated list of ages or "", Split will give an empty array if this property is ever Empty
                aryChildAges = _.SPLIT(_.CALLm1v0(this, objUnit, "ChildrenAges"), ",");
                var loopEnd27 = _.UBOUND(aryChildAges);
                var loopStart27 = _.NUM((Int16)0, loopEnd27, (Int16)1);
                if (_.StrictLTE(loopStart27, loopEnd27))
                {
                    for (iChildAgeIndex = loopStart27; _.StrictLTE(iChildAgeIndex, loopEnd27); iChildAgeIndex = _.ADD(iChildAgeIndex, (Int16)1))
                    {
                        _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"hidden\" name=\"roomReq_", aiThisReqmnt, "_children_childage", iChildAgeIndex, "\" value=\"", _.CALLm0argp(this, aryChildAges, _.ARGS.Ref(iChildAgeIndex, v205 => { iChildAgeIndex = v205; })), "\" />"));
                    }
                }

            }

            _.CALLm1v1(this, pO, "Write", _.CONCAT("</div>", VBScriptConstants.vbCrLf));
            _.CALLm1v1(this, pO, "Write", _.CONCAT("<div class=\"pnStayReqmntRslts\">", VBScriptConstants.vbCrLf));

            return BookingUI_RenderNewReq_retVal;
        }
        // SUMMARY: render unit option HTML
        // [aiStayNum]: integer stay index
        // [aiThisReqmnt]: integer requirement index
        // [aiUnitKey]: integer unit key
        // [bSelected]: should the current unit appear selected
        // [arsAvail]: ADO availability recordset
        // [asAvailClassName]: string avail class name
        public object BookingUI_RenderUnit(object aiStayNum, object aiThisReqmnt, object bSelected, object objAvailEntry, object objUnit, object objAllUnits, object asAvailClassName, object pO, object bRenderMaximumUnitsAvailable)
        {
            object BookingUI_RenderUnit_retVal = null;
            object mUnitStayTotal = null;
            object iNumNights = null;
            object mUnitPerNight = null;
            object iNumPeople = null;
            object iDaysBreakfast = null;
            object bPerPerson = null;
            object mPersonPerNight = null;
            object strIptId = null;
            object mUnitStayTotalPayableBasedOnGuidePrice = null;
            object bDiscountApplied = null;
            object iAdults = null;
            object iChildren = null;
            object iMaxUnitsAvailable = null;
            object UnitCostPerPerson = null; /* Undeclared in source */

            mUnitStayTotal = _.VAL(_.CALLm1v0(this, objUnit, "StayTotalPayable"));
            mUnitStayTotalPayableBasedOnGuidePrice = _.VAL(_.CALLm1v0(this, objUnit, "StayTotalPayableBasedOnGuidePrice"));
            iNumNights = _.VAL(_.CALLm1v0(this, objAvailEntry, "Nights"));
            mUnitPerNight = _.DIV(mUnitStayTotal, iNumNights);
            bPerPerson = _.VAL(_.CALLm1v0(this, objUnit, "Perperson"));
            iNumPeople = _.VAL(_.CALLm1v0(this, objUnit, "ReqSize"));

            iDaysBreakfast = _.VAL(_.CALLm1v0(this, objUnit, "DaysBreakfast"));
            bDiscountApplied = _.VAL(_.CALLm1v0(this, objUnit, "IncludesChildDiscount"));

            iMaxUnitsAvailable = _.VAL(_.CALLm1v0(this, objUnit, "MaximumQuantityAvailable"));

            // We need an id so we can set the label's "for" attribute, but if VB Polling is enabled,
            // we might end up with id duplication - so in that case we append a random suffix
            strIptId = _.CONCAT("unit_", aiStayNum, "_", aiThisReqmnt, "_", _.CALLm1v0(this, objUnit, "UnitKey"));
            if (_.IF(_outer.IsVBPollingEnabled))
            {
                strIptId = _.CONCAT(strIptId, "_", _.INT(_.MULT(_.RND(), 100000)));
            }

            _.CALLm1v1(this, pO, "Write", "<div class=\"pnUnitOption\">");
            _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"radio\" name=\"unit_", aiStayNum, "_", aiThisReqmnt, "\" ", "id=\"", strIptId, "\" "));
            if (_.IF(_.NOT(_outer.IsVBPollingEnabled)))
            {
                // Not sure this onclick is even required without VB Polling.. (?)
                _.CALLm1v1(this, pO, "Write", "onclick=\"BookingUI_UnitSelect(this);\" ");
            }
            _.CALLm1v1(this, pO, "Write", _.CONCAT("value=\"", _.CALLm1v0(this, objUnit, "UnitKey"), "\" "));
            if (_.IF(bSelected))
            {
                _.CALLm1v1(this, pO, "Write", "checked=\"checked\" ");
            }
            _.CALLm1v1(this, pO, "Write", "/>");
            _.CALLm1v1(this, pO, "Write", _.CONCAT("<label for=\"", strIptId, "\"> "));
            _.CALLm1v1(this, pO, "Write", _.CONCAT(_.CALLm1v0(this, objUnit, "UnitName"), " - ", _.CALLm1argp(this, _outer, "BookingUI_NicePrice", _.ARGS.Ref(mUnitStayTotal, v207 => { mUnitStayTotal = v207; })), " ", asAvailClassName));

            //if we have child pricing discount applied show the icon
            if (_.IF(bDiscountApplied))
            {
                _.CALLm1v1(this, pO, "Write", _.CALLm1v1(this, _outer, "BookingUI_AvailClassIcon", "DISCOUNT"));
            }

            _.CALLm1v1(this, pO, "Write", "</label>");
            _.CALLm1v1(this, pO, "Write", _.CONCAT("</div>", VBScriptConstants.vbCrLf));
            _.CALLm1v1(this, pO, "Write", _.CONCAT("<div class=\"pnPriceBase\">", VBScriptConstants.vbCrLf));

            //#MJ 29/04/2010 -	decision made not to show the price basis as the per person figure was always a guestimate, child pricing messes with the price so per person doesn't apply
            //					also we now always deal with total stay prices
            //				If bPerPerson Then
            //					mPersonPerNight = mUnitPerNight/iNumPeople
            //					pO.Write BookingUI_NicePrice(mPersonPerNight) & " " & Page.Resource("bookonline/unitselection/perpersonpernight", "per person per night") & ". "
            //				Else
            //					pO.Write BookingUI_NicePrice(mUnitPerNight) & " " & Page.Resource("bookonline/unitselection/perroomunitpernight", "per room/unit per night") & ". "
            //				End If

            if (_.IF(_.EQ(iDaysBreakfast, iNumNights)))
            {
                _.CALLm1v1(this, pO, "Write", _.CONCAT(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/breakfastincluded", "Breakfast included"), ". "));
            }
            else if (_.IF(_.GT(_.NullableNUM(iDaysBreakfast), (Int16)0)))
            {
                _.CALLm1v1(this, pO, "Write", _.CONCAT(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/breakfastincludedon", "Breakfast included on "), iDaysBreakfast, " ", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/day(s)", "day(s)"), ". "));
            }

            if (_.IF(_.LT(iNumPeople, _.CALLm1v0(this, objUnit, "MinOcc"))))
            {
                if (_.IF(bPerPerson))
                {
                    _.CALLm1v1(this, pO, "Write", "<br />");
                    _.CALLm1v1(this, pO, "Write", _.CONCAT(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/priceperpersonincludes", "Price Per Person includes"), " "));
                    _.CALLm1v1(this, pO, "Write", _.CALLm1v1(this, _outer, "BookingUI_NicePrice", _.SUBT(mPersonPerNight, _.DIV(UnitCostPerPerson, iNumNights))));
                    _.CALLm1v1(this, pO, "Write", _.CONCAT(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/minimumoccupancysupplement", " minimum occupancy supplement"), ". "));
                }
                else
                {
                    _.CALLm1v1(this, pO, "Write", _.CONCAT(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/minoccupancyof", "Min. occupancy of"), " ", _.CALLm1v0(this, objUnit, "MinOcc"), ". "));
                }
            }
            _.CALLm1v1(this, pO, "Write", _.CONCAT("<div class=\"pnLinkedUnit\">", _.CALLm1argp(this, _outer, "BookingUI_LinkedUnitDesc", _.ARGS.Ref(objUnit, v209 => { objUnit = v209; }).Ref(objAllUnits, v210 => { objAllUnits = v210; })), "</div>"));

            _.CALLm1v1(this, pO, "Write", _.CONCAT("</div>", VBScriptConstants.vbCrLf));

            if (_.IF(_.NOT(_.CALLm1v0(this, objAvailEntry, "IsLocal"))))
            {
                _.CALLm1v1(this, pO, "Write", "<input type=\"hidden\" ");
                _.CALLm1v1(this, pO, "Write", _.CONCAT("name=\"uxml_", aiStayNum, "_", aiThisReqmnt, "_", _.CALLm1v0(this, objUnit, "UnitKey"), "\" "));
                _.CALLm1v1(this, pO, "Write", _.CONCAT("value=\"", _.CALLm1v1(this, _outer.Server, "HtmlEncode", _.CALLm1v0(this, objUnit, "EviivoMetaData")), "\" />"));
            }

            if (_.IF(bRenderMaximumUnitsAvailable))
            {
                _.CALLm1v1(this, pO, "Write", "<div class=\"maxAvailUnits\">");
                _.CALLm1v1(this, pO, "Write", "<p>");
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<span class=\"maxAvailUnitsLabelPrefix\">", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/maxiumunitsavailableprefix", "Only "), "</span>"));
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<span class=\"maxAvailUnitsValue\">", _.CALLm1v0(this, objUnit, "MaximumQuantityAvailable"), "</span>"));
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<span class=\"maxAvailUnitsLabelSuffix\">", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/maxiumunitsavailablesuffix", " Rooms Remaining"), "</span>"));
                _.CALLm1v1(this, pO, "Write", "</p>");
                _.CALLm1v1(this, pO, "Write", "</div>");
            }

            return BookingUI_RenderUnit_retVal;
        }
        // ====================================================================================================
        // RENDER: Generate markup for booking buttons (only used by Acco, not Ticketing)
        // ====================================================================================================
        // SUMMARY: render BOOK and BACK buttons
        // [aiStayNum]: integer stay number [1-based]
        // [abPrecise]: boolean precise match (ie. hide BACK button)
        // <retval>: string output
        public object BookingUI_RenderButtons(object aiStayNum, object pO, object bExternal)
        {
            object BookingUI_RenderButtons_retVal = null;
            object strClass = null;
            strClass = "btnBookStay";

            if (_.IF(bExternal))
            {
                strClass = _.CONCAT(strClass, " redirect");
            }

            _.CALLm1v1(this, pO, "Write", "<div class=\"pnStayButtons\">");
            _.CALLm1v1(this, pO, "Write", "<input ");
            _.CALLm1v1(this, pO, "Write", "type=\"image\" ");
            _.CALLm1v1(this, pO, "Write", _.CONCAT("class=\"", strClass, "\" "));
            _.CALLm1v1(this, pO, "Write", _.CONCAT("name=\"bookstay_", aiStayNum, "\" "));

            // Not using ids with VB Polling layout
            if (_.IF(_.NOT(_outer.IsVBPollingEnabled)))
            {
                _.CALLm1v1(this, pO, "Write", _.CONCAT("id=\"bookstay_", aiStayNum, "\" "));
            }

            _.CALLm1v1(this, pO, "Write", _.CONCAT("value=\"", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/btn/book", "Book"), "\" "));
            _.CALLm1v1(this, pO, "Write", _.CONCAT("src=\"", _.CALLm1v2(this, _outer.Page, "ImageResource", "bookonline/btn/book", _.CONCAT(_.CALLm1v0(this, _outer.Context, "ImageDir"), "booking/book.gif")), "\" "));
            _.CALLm1v1(this, pO, "Write", _.CONCAT("alt=\"", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/btn/book", "Book"), "\" />"));

            _.CALLm1v1(this, pO, "Write", _.CONCAT("</div>", VBScriptConstants.vbCrLf));

            return BookingUI_RenderButtons_retVal;
        }
        // ====================================================================================================
        // RENDER: Translate availability options -> action description string
        //  eg. Telephone Booking       -> "Submit a Booking Enquiry"
        //      Indicative Availability -> "Confirm availability"
        // ====================================================================================================
        // SUMMARY: describe the significant availability type (eg. Indicative, etc)
        // [abIndicative]: boolean indicates whether there is INDICATIVE availability
        // [abIndicValid]: boolean is INDICATIVE availability valid in this case
        // [abTeleBook]: boolean on telebook channel
        public object BookingUI_AvailClassName(object abIndicative, object abIndicValid, object abTeleBook)
        {
            object BookingUI_AvailClassName_retVal = null;
            // If telephone booking, there's only one option
            if (_.IF(abTeleBook))
            {
                BookingUI_AvailClassName_retVal = _.VAL(_.CALLm1v1(this, _outer, "BookingUI_AvailClassIcon", "TELE"));
                return BookingUI_AvailClassName_retVal;
            }

            // If not telephone and not indicative, must be allocated
            if (_.IF(_.NOT(abIndicative)))
            {
                BookingUI_AvailClassName_retVal = _.VAL(_.CALLm1v1(this, _outer, "BookingUI_AvailClassIcon", "ALLOC"));
                return BookingUI_AvailClassName_retVal;
            }

            // Otherwise, get appropriate indicative option
            if (_.IF(abIndicValid))
            {
                BookingUI_AvailClassName_retVal = _.VAL(_.CALLm1v1(this, _outer, "BookingUI_AvailClassIcon", "INDIC"));
            }
            else
            {
                BookingUI_AvailClassName_retVal = _.VAL(_.CALLm1v1(this, _outer, "BookingUI_AvailClassIcon", "TELE"));
            }
            return BookingUI_AvailClassName_retVal;
        }
        // ====================================================================================================
        // RENDER:Translate avail class ID -> action description string
        //  eg. "ALLOC" -> "Book Online"
        //      "INDIC" -> "Confirm Availability"
        // ====================================================================================================
        // SUMMARY: get HTML for rendering the icon describing the availability class
        // [asAvailClassId]: string availability class ID (ie. ALLOC, INDIC or TELE)
        // <retval>: string image describing availClass
        public object BookingUI_AvailClassIcon(ref object asAvailClassId)
        {
            object BookingUI_AvailClassIcon_retVal = null;
            object sIcon = null;
            object sTxt = null;
            object sImg = null;

            //Select Case asAvailClassId
            //	Case "ALLOC"
            //		sIcon="icon_availClass_alloc"
            //		sTxt=Page.Resource("bookonline/btn/bookonline","Book Online")
            //	Case "INDIC"
            //		sIcon="icon_availClass_indic"
            //		sTxt=Page.Resource("bookonline/btn/confirmavailability","Confirm Availability")
            //	Case "DISCOUNT"
            //		sIcon="icon_availClass_discount"
            //		sTxt=Page.Resource("bookonline/btn/discountapplied","Child Pricing Discount Applied")
            //	Case Else
            //		sIcon="icon_availClass_tele"
            //		sTxt=Page.Resource("bookonline/btn/submitbookingenquiry","Submit a Booking Enquiry")
            //End Select

            sImg = _.VAL(_.CALLm1v2(this, _outer.Page, "ImageResource", _.CONCAT("bookonline/icons/", sIcon), _.CONCAT(_.CALLm1v0(this, _outer.Context, "ImageDir"), "booking/", sIcon, ".gif")));
            BookingUI_AvailClassIcon_retVal = _.CONCAT("<img src=\"", sImg, "\" style=\"vertical-align:middle;\" alt=\"", sTxt, "\" />");
            return BookingUI_AvailClassIcon_retVal;
        }
        // ====================================================================================================
        // RENDER: Format currency value
        // ====================================================================================================
        // SUMMARY: save space - only display price with pennies digits when fractional pounds
        // [amPrice]: money price to render
        // <retval>: string price
        public object BookingUI_NicePrice(ref object amPrice)
        {
            object BookingUI_NicePrice_retVal = null;
            object strPrice = null;
            // Get price:
            // - MakePrice will also handle any currency conversion)
            // - MakePrice will apply an appropriate currency symbol
            object byrefalias40 = amPrice;
            try
            {
                strPrice = _.VAL(_.CALLm3argp(this, _outer.Page, "Functions", "Money", "MakePrice", _.ARGS.Ref(byrefalias40, v213 => { byrefalias40 = v213; })));
            }
            finally { amPrice = byrefalias40; }

            // If there's a trailing ".00" then trim it off
            // NB: Pretty sure we'll never get a price of the form "?.00" - it should always
            //     be "?0.00", but just in case check that we've got a suitable long string
            if (_.IF(_.GT(_.NullableNUM(_.LEN(strPrice)), (Int16)4)))
            {
                if (_.IF(_.EQ(_.NullableSTR(_.RIGHT(strPrice, (Int16)3)), ".00")))
                {
                    strPrice = _.VAL(_.LEFT(strPrice, _.SUBT(_.LEN(strPrice), (Int16)3)));
                }
            }

            // Return string ready for display
            BookingUI_NicePrice_retVal = _.VAL(_.CALLm1argp(this, _outer.Server, "HTMLEncode", _.ARGS.Ref(strPrice, v214 => { strPrice = v214; })));
            return BookingUI_NicePrice_retVal;
        }
        // ====================================================================================================
        // RENDER: Pull description of linked unit (includes name of linked unit, name of source unit and
        // size of linked unit)
        // ====================================================================================================
        // SUMMARY: get description of linked unit - this is the PHYSICAL unit description
        public object BookingUI_LinkedUnitDesc(object objUnit, object objAllUnits)
        {
            object BookingUI_LinkedUnitDesc_retVal = null;
            object sUnitName = null;
            object sLinkedUnitName = null;
            object objParentUnit = null;
            object intIndex = null;

            // If either UnitName of LinkedUnitName absent, return blank
            sUnitName = _.VAL(_.CALLm1v0(this, objUnit, "UnitName"));
            sLinkedUnitName = _.VAL(_.CALLm1v0(this, objUnit, "LinkUnitName"));
            if (_.IF(_.OR(_.OR(_.OR(_.ISNULL(sUnitName), _.EQ(_.NullableSTR(sUnitName), "")), _.ISNULL(sLinkedUnitName)), _.EQ(_.NullableSTR(sLinkedUnitName), ""))))
            {
                BookingUI_LinkedUnitDesc_retVal = "";
                return BookingUI_LinkedUnitDesc_retVal;
            }

            // 2014-08-26 DWR: We need to retrieve the capacity of the unit that this linked unit is linked to. This data is not available in the avail
            // data from TOv2 since it is not included in the data from the Availability Component. It is why the "all units" data must be passed into
            // this method. This change addresses FogBugz 12998.
            objParentUnit = VBScriptConstants.Nothing;
            var loopEnd28 = _.NUM(_.SUBT(_.CALLm1v0(this, objAllUnits, "Count"), (Int16)1));
            var loopStart28 = _.NUM((Int16)0, loopEnd28, (Int16)1);
            if (_.StrictLTE(loopStart28, loopEnd28))
            {
                for (intIndex = loopStart28; _.StrictLTE(intIndex, loopEnd28); intIndex = _.ADD(intIndex, (Int16)1))
                {
                    if (_.IF(_.EQ(_.CALLm1v0(this, _.CALLm1argp(this, objAllUnits, "getItem", _.ARGS.Ref(intIndex, v215 => { intIndex = v215; })), "Key"), _.CALLm1v0(this, objUnit, "LinkUnitKey"))))
                    {
                        objParentUnit = _.OBJ(_.CALLm1argp(this, objAllUnits, "getItem", _.ARGS.Ref(intIndex, v216 => { intIndex = v216; })));
                        break;
                    }
                }
            }
            if (_.IF(_.IS(objParentUnit, VBScriptConstants.Nothing)))
            {
                _.CALLm1v1(this, _outer.Page, "PrintTraceWarning", _.CONCAT("Unable to locate parent unit (", _.CALLm1v0(this, objUnit, "LinkUnitKey"), ") for linked unit ", _.CALLm1v0(this, objUnit, "UnitKey")));
                BookingUI_LinkedUnitDesc_retVal = "";
                return BookingUI_LinkedUnitDesc_retVal;
            }

            BookingUI_LinkedUnitDesc_retVal = _.REPLACE(_.REPLACE(_.REPLACE(_.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/alsosoldaswithpersoncapacity", "(<i>#linkedunitname#</i> sold as #unitname# with #linkunitsize# person capacity)"), "#linkedunitname#", sLinkedUnitName), "#unitname#", sUnitName), "#linkunitsize#", _.CALLm1v0(this, objParentUnit, "Capacity"));
            return BookingUI_LinkedUnitDesc_retVal;
        }
        // ====================================================================================================
        // RENDER: This handles all of the rendering for ticketing - none of the StaySummary, StayDetails,
        // RenderButtons malarkey is required
        // ====================================================================================================
        public object BookingUI_TicketsSummary(ref object objAvailEntry, ref object adtStartNight, ref object pO)
        {
            object BookingUI_TicketsSummary_retVal = null;
            object iTotal = null;
            object iSubTotal = null;
            object iSelectedQty = null;
            object intIndexUnit = null;
            object objUnit = null;
            object strPriceBasis = null;

            if (_.IF(_.GT(_.NullableNUM(_.CALLm2v0(this, objAvailEntry, "Units", "Count")), (Int16)0)))
            {
                _.CALLm1v1(this, pO, "Write", "<div id=\"availabilityCalendarTableWrapper\">");
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<h3>", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/ticketsavailable", "Tickets Available:"), "</h3>"));
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<table id=\"availabilityCalendarTable\" summary=\"", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/ticketsavailable", "Tickets Available"), "\" border=\"1\">"));
                _.CALLm1v1(this, pO, "Write", "<thead>");
                _.CALLm1v1(this, pO, "Write", "<tr class=\"heading\">");
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<th class=\"unit\">", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/tickets", "Tickets"), "</th>"));
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<th class=\"select\">", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/selection", "Selection"), "</th>"));
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<th class=\"date\">", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/date", "Date"), "</th>"));
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<th class=\"total\">", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/total", "Total"), "</th>"));
                _.CALLm1v1(this, pO, "Write", "</tr>");
                _.CALLm1v1(this, pO, "Write", "<tr>");
                _.CALLm1v1(this, pO, "Write", "<th></th>");
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<th class=\"number\">", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/nooftickets", "No.Tickets"), "</th>"));
                object byrefalias41 = adtStartNight;
                try
                {
                    _.CALLm1v1(this, pO, "Write", _.CONCAT("<th class=\"staydate\">", _.CALLm3argp(this, _outer.Page, "Functions", "Dates", "NiceDateGuts", _.ARGS.Ref(byrefalias41, v217 => { byrefalias41 = v217; }).Val(true).Val(true)), "</th>"));
                }
                finally { adtStartNight = byrefalias41; }
                _.CALLm1v1(this, pO, "Write", "<th class=\"total\"></th>");
                _.CALLm1v1(this, pO, "Write", "</tr>");
                _.CALLm1v1(this, pO, "Write", "</thead>");
                _.CALLm1v1(this, pO, "Write", "<tbody>");
                iTotal = (Int16)0;

                var loopEnd29 = _.NUM(_.SUBT(_.CALLm2v0(this, objAvailEntry, "Units", "Count"), (Int16)1));
                var loopStart29 = _.NUM((Int16)0, loopEnd29, (Int16)1);
                if (_.StrictLTE(loopStart29, loopEnd29))
                {
                    for (intIndexUnit = loopStart29; _.StrictLTE(intIndexUnit, loopEnd29); intIndexUnit = _.ADD(intIndexUnit, (Int16)1))
                    {
                        objUnit = _.OBJ(_.CALLm2argp(this, objAvailEntry, "Units", "GetItem", _.ARGS.Ref(intIndexUnit, v219 => { intIndexUnit = v219; })));

                        iSelectedQty = _.CLNG(_.CALLm1v1(this, _outer.Request, "Form", _.CONCAT("unit_", _.CALLm1v0(this, objUnit, "UnitKey"))));

                        if (_.IF(_.CALLm1v0(this, objUnit, "PerPerson")))
                        {
                            strPriceBasis = "per per";
                        }
                        else
                        {
                            strPriceBasis = "per tic";
                        }

                        _.CALLm1v1(this, pO, "Write", _.CONCAT("<tr id=\"row_", _.CALLm1v0(this, objUnit, "UnitKey"), "\">"));
                        _.CALLm1v1(this, pO, "Write", _.CONCAT("<td class=\"unit\">", _.CALLm1v0(this, objUnit, "UnitName"), "</td>"));
                        _.CALLm1v1(this, pO, "Write", _.CONCAT("<td class=\"select\">", _.CALLm3argp(this, _outer.Page, "Functions", "Booking", "DrawSelectRange", _.ARGS.Val(_.CONCAT("unit_", _.CALLm1v0(this, objUnit, "UnitKey"))).Val((Int16)0).Val(_.CALLm1v0(this, objUnit, "UnitCount")).Ref(iSelectedQty, v220 => { iSelectedQty = v220; })), "</td>"));
                        _.CALLm1v1(this, pO, "Write", _.CONCAT("<td class=\"price\">", _.CALLm1v1(this, _outer.Server, "HTMLEncode", _.CALLm3v1(this, _outer.Page, "Functions", "Money", "MakePrice", _.CALLm1v0(this, objUnit, "StayTotalPayable"))), "</td>"));
                        _.CALLm1v1(this, pO, "Write", _.CONCAT("<td class=\"total\">", "<input type=\"hidden\" name=\"data_", _.CALLm1v0(this, objUnit, "UnitKey"), "\" id=\"data_", _.CALLm1v0(this, objUnit, "UnitKey"), "\" value=\"", _.CALLm1v0(this, objUnit, "UnitCount"), ",", _.CALLm1v0(this, objUnit, "MinOcc"), ",", _.CALLm1v0(this, objUnit, "UnitSize"), ",", strPriceBasis, ",", _.CALLm1v0(this, objUnit, "StayTotalPayable"), "\">", _.CALLm1v1(this, _outer.Server, "HTMLEncode", _.CALLm3v1(this, _outer.Page, "Functions", "Money", "MakePrice", _.MULT(_.CALLm1v0(this, objUnit, "StayTotalPayable"), iSelectedQty))), "</td>"));
                        _.CALLm1v1(this, pO, "Write", "</tr>");
                        iTotal = _.ADD(iTotal, _.MULT(_.CALLm1v0(this, objUnit, "StayTotalPayable"), iSelectedQty));

                    }
                }
                iSubTotal = _.ADD(iSubTotal, iTotal);

                _.CALLm1v1(this, pO, "Write", "</tbody>");
                _.CALLm1v1(this, pO, "Write", "</table>");
                _.CALLm1v1(this, pO, "Write", "</div>");
                _.CALLm1v1(this, pO, "Write", "<table id=\"availabilityTotals\" summary=\"Totals\" border=\"1\">");
                _.CALLm1v1(this, pO, "Write", "<tr>");
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<th>", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/unitselection/grandtotal", "Grand Total"), "</th>"));
                _.CALLm1v1(this, pO, "Write", "<noscript>");
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<td><input type=\"image\" src=\"", _.CALLm1v2(this, _outer.Page, "ImageResource", "bookonline/unitselection/recalculate", _.CONCAT(_.CALLm1v0(this, _outer.Context, "ImageDir"), "booking/bookrecalculate.gif")), "\" name=\"recalculate\" value=\"recalculate\" class=\"submit\"/></td>"));
                _.CALLm1v1(this, pO, "Write", "</noscript>");
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<td id=\"AvCalTotal\">", _.CALLm1v1(this, _outer.Server, "HTMLEncode", _.CALLm3argp(this, _outer.Page, "Functions", "Money", "MakePrice", _.ARGS.Ref(iSubTotal, v222 => { iSubTotal = v222; }))), "</td>"));
                _.CALLm1v1(this, pO, "Write", "</tr>");
                _.CALLm1v1(this, pO, "Write", "</table>");
                _.CALLm1v1(this, pO, "Write", _.CONCAT("<input type=\"image\" src=\"", _.CALLm1v2(this, _outer.Page, "ImageResource", "bookonline/btn/bookticketing", _.CONCAT(_.CALLm1v0(this, _outer.Context, "ImageDir"), "booking/bookticketing.gif")), "\" name=\"bookit\" value=\"", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/btn/book", "Book"), "\" alt=\"", _.CALLm1v2(this, _outer.Page, "Resource", "bookonline/btn/book", "Book"), "\" class=\"submit\"/>"));
            }
            return BookingUI_TicketsSummary_retVal;
        }
        // ====================================================================================================
        // RENDER: If Site Param Booking_ForceExternal is set, then all  bookings are forced to external sites
        // (the site depends upon the product's Estate) - this function gets the external destination
        // 2009-02-19 DWR: This handles Local availability products, FrontDesk products will be treated like
        // this as well if VB Polling is disabled, if it is ENabled then the FrontDesk products will act like
        // any other external supplier and should deep-link into their site.
        // ====================================================================================================
        public object GetExtBookUrlFromProductEstate(ref object asEstateID)
        {
            object GetExtBookUrlFromProductEstate_retVal = null;
            object strPostUrl_Ext = null;
            object strPostUrl_ExtDflt = null;
            object aryExtBookEstate = null;
            object i = null;
            // 2009-02-13 DWR: Can't remove spaces from content here because estate ids can have
            // spaces in (eg. "Arun DC" in TSE)
            aryExtBookEstate = _.SPLIT(_.REPLACE(_.TRIM(_.CONCAT("", _.CALLm2v1(this, _outer.Page, "Site", "Params", "Booking_ExtBookEstateMapping"))), VBScriptConstants.vbCrLf, ""), ",");
            var loopEnd30 = _.NUM(_.SUBT(_.UBOUND(aryExtBookEstate), (Int16)1));
            var loopStart30 = _.NUM((Int16)0, loopEnd30, (Int16)2);
            if (_.StrictLTE(loopStart30, loopEnd30))
            {
                for (i = loopStart30; _.StrictLTE(i, loopEnd30); i = _.ADD(i, (Int16)2))
                {
                    if (_.IF(_.EQ(_.NullableSTR(_.UCASE(_.TRIM(_.CALLm0argp(this, aryExtBookEstate, _.ARGS.Ref(i, v224 => { i = v224; }))))), "DEFAULT")))
                    {
                        strPostUrl_ExtDflt = _.VAL(_.CALLm0argp(this, aryExtBookEstate, _.ARGS.Val(_.ADD(i, (Int16)1))));
                    }
                    else if (_.IF(_.EQ(_.UCASE(_.TRIM(_.CALLm0argp(this, aryExtBookEstate, _.ARGS.Ref(i, v225 => { i = v225; })))), _.UCASE(_.TRIM(asEstateID)))))
                    {
                        strPostUrl_Ext = _.VAL(_.CALLm0argp(this, aryExtBookEstate, _.ARGS.Val(_.ADD(i, (Int16)1))));
                        break;
                    }
                }
            }

            if (_.IF(_.NOTEQ(_.NullableSTR(strPostUrl_Ext), "")))
            {
                GetExtBookUrlFromProductEstate_retVal = _.VAL(strPostUrl_Ext);
                _.CALLm1v1(this, _outer.Page, "PrintTrace", _.CONCAT("GetExtBookUrlFromProductEstate: Product Estate ID = ", asEstateID, ", External Book Url = ", strPostUrl_Ext));
            }
            else if (_.IF(_.NOTEQ(_.NullableSTR(strPostUrl_ExtDflt), "")))
            {
                GetExtBookUrlFromProductEstate_retVal = _.VAL(strPostUrl_ExtDflt);
                _.CALLm1v1(this, _outer.Page, "PrintTrace", _.CONCAT("GetExtBookUrlFromProductEstate: Product Estate ID = ", asEstateID, ", Using Default External Book Url = ", strPostUrl_ExtDflt));
            }
            else
            {
                _.RAISEERROR(_.ADD(VBScriptConstants.vbObjectError, (Int16)1), "ETWP.Booking_UnitSelection Control", _.CONCAT("Failed to get External Booking Url [", asEstateID, "]"));
            }

            return GetExtBookUrlFromProductEstate_retVal;
        }
        public object InitExternalBookingSettings()
        {
            object InitExternalBookingSettings_retVal = null;
            if (_.IF(_.AND(_.CALLm2v1(this, _outer.Page, "Site", "Params", "Booking_ForceExternal"), _.NOTEQ(_.NullableSTR(_.TRIM(_.CONCAT("", _.CALLm2v1(this, _outer.Page, "Site", "Params", "Booking_ExtBookEstateMapping")))), ""))))
            {
                _outer.IsExternalBooking = true;
            }
            else
            {
                _outer.IsExternalBooking = false;
            }
            return InitExternalBookingSettings_retVal;
        }
        // ====================================================================================================
        // MISC: Since the RenderSettings.BookingRequirement references passed into here are usually read-only
        // instances from the Page.Functions.GetSharedObject method, we'll need to make a local copy that we
        // can manipulate (since in some cases we need to mess about with the values)
        // ====================================================================================================
        public object GetEditableBookingRequirement(object objBookingRequirement)
        {
            object GetEditableBookingRequirement_retVal = null;
            object objBookingRequirementNew = null;

            objBookingRequirementNew = _.OBJ(_.CALLm2v1(this, _outer.Page, "Functions", "GetNewObject", "BookingRequirement"));
            var with = _.OBJ(objBookingRequirementNew);
            _.SETm1a0(this, with, "VisitDate", _.VAL(_.CALLm1v0(this, objBookingRequirement, "VisitDate")));
            _.SETm1a0(this, with, "Nights", _.VAL(_.CALLm1v0(this, objBookingRequirement, "Nights")));
            _.SETm1a0(this, with, "FlexibleRange", _.VAL(_.CALLm1v0(this, objBookingRequirement, "FlexibleRange")));
            _.SETm1a0(this, with, "Adults", _.VAL(_.CALLm1v0(this, objBookingRequirement, "Adults")));
            _.SETm1a0(this, with, "Children", _.VAL(_.CALLm1v0(this, objBookingRequirement, "Children")));
            _.SETm1a0(this, with, "ChildAges", _.VAL(_.CALLm1v0(this, objBookingRequirement, "ChildAges")));
            _.SETm1a0(this, with, "IsEviivoBooking", _.VAL(_.CALLm1v0(this, objBookingRequirement, "IsEviivoBooking")));
            _.SETm1a0(this, with, "Consumer", _.VAL(_.CALLm1v0(this, objBookingRequirement, "Consumer")));
            _.SETm1a0(this, with, "Offer", _.VAL(_.CALLm1v0(this, objBookingRequirement, "Offer")));
            _.SETm1a0(this, with, "BookingPassword", _.VAL(_.CALLm1v0(this, objBookingRequirement, "BookingPassword")));
            _.SETm1a0(this, with, "Product", _.VAL(_.CALLm1v0(this, objBookingRequirement, "Product")));
            _.SETm1a0(this, with, "Requirement", _.VAL(_.CALLm1v0(this, objBookingRequirement, "Requirement")));
            _.SETm1a0(this, with, "RequirementRef", _.VAL(_.CALLm1v0(this, objBookingRequirement, "RequirementRef")));
            // NP 2012-03-12: RoomRequirements are needed
            // See GenerateRequirementFormData and Page.Functions.Booking.GenerateRequirementKeyValueData
            // the "NumRoomReq" value is part of the RoomRequirement, if it is not available then GenerateRequirementKeyValueData
            // sets default values for the adult and number of room requirements both to 1.
            // Requirements are not being passed to the RequirementSummary control correctly because the BookingRequestDictionary
            // is being overwritten with these incorrect default values.
            _.SETm1a0(this, with, "RoomRequirements", _.VAL(_.CALLm1v0(this, objBookingRequirement, "RoomRequirements")));
            GetEditableBookingRequirement_retVal = _.OBJ(objBookingRequirementNew);
            return GetEditableBookingRequirement_retVal;
        }
    }
    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object hlContext { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}
