using System;
using System.Collections;
using System.Runtime.InteropServices;
using Skrypton.RuntimeSupport;
using Skrypton.RuntimeSupport.Attributes;
using Skrypton.RuntimeSupport.Exceptions;
using Skrypton.RuntimeSupport.Compat;

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

            _outer.booking_pollingredirect = (Int16)3;
            _outer.booking_redirect = (Int16)2;
            _outer.booking_eviivo = (Int16)1;
            _outer.booking_local = (Int16)0;

            _outer.interfaceversion = (Int16)1;

            //nasty globals
            _outer.g_inumberofcalendarsrendered = (Int16)0;
            _outer.bformrendered = false;

            _outer.bprodhasavail = false;

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
            page = null;
            request = null;
            context = null;
            server = null;
            dms = null;
            interfaceversion = null;
            booking_local = null;
            booking_eviivo = null;
            booking_redirect = null;
            booking_pollingredirect = null;
            isexternalbooking = null;
            strextbookurl = null;
            strproductestateid = null;
            bformrendered = null;
            isvbpollingenabled = null;
            brenderascalendar = null;
            g_inumberofcalendarsrendered = null;
            bprodhasavail = null;
        }

        internal object page { get; set; }
        internal object request { get; set; }
        internal object context { get; set; }
        internal object server { get; set; }
        internal object dms { get; set; }
        internal object interfaceversion { get; set; }
        internal object booking_local { get; set; }
        internal object booking_eviivo { get; set; }
        internal object booking_redirect { get; set; }
        internal object booking_pollingredirect { get; set; }
        internal object isexternalbooking { get; set; }
        internal object strextbookurl { get; set; }
        internal object strproductestateid { get; set; }
        internal object bformrendered { get; set; }
        internal object isvbpollingenabled { get; set; }
        internal object brenderascalendar { get; set; }
        internal object g_inumberofcalendarsrendered { get; set; }
        internal object bprodhasavail { get; set; }

        public object getprodhasavail()
        {
            return _.VAL(_outer.bprodhasavail);
        }

        // ====================================================================================================
        // RENDER: Availability Calendar (supports local availability only!)
        // - Note: This doesn't actually perform any data access, all of the content required is passed
        //   through in POST data from the availability calendar on the previous page (Product Detail)
        // ====================================================================================================
        public object bookingui_staymain_availcal(object po, object objrendersettings)
        {
            object BookingUI_StayMain_AvailCal_retVal = null;
            object objbookingrequirement = null;
            object intbookingtype = null;
            object istaynum = null;
            object ithisreqmnt = null;
            object iunitqty = null;
            object iunitminoccupancy = null;
            object iunitmaxcapacity = null;
            object iunitkey = null;
            object ilinkedunitkey = null;
            object strunitname = null;
            object stravailclassid = null;
            object item = null;
            object strtemp = null;
            object i = null;
            object dstart = null;
            object inights = null;
            object intprodkey = null;
            // Expect selections as set of form values:
            //  "unit_prodkey", "minoccu_prodkey", "maxcap_prodkey", "name_prodkey", "availclass_prodkey"
            //
            // If linked units are referenced, the first value will be:
            //  "unit_prodkey_linkprodkey"

            // 2011-08-09 DWR: Get populated read-only Booking Requirement data From GetSharedObject, then translate into a local copy we can edit
            // (since some methods in here try to mess about with properties on it)
            objbookingrequirement = _.OBJ(_.CALL(this, _outer.page, "Functions", "GetSharedObject", _.ARGS.Val("BookingRequirement")));
            objbookingrequirement = _.OBJ(_.CALL(this, _outer, "GetEditableBookingRequirement", _.ARGS.Ref(objbookingrequirement, v => { objbookingrequirement = v; })));

            dstart = _.VAL(_.CALL(this, objbookingrequirement, "VisitDate"));
            inights = _.VAL(_.CALL(this, objbookingrequirement, "Nights"));
            intprodkey = _.VAL(_.CALL(this, objbookingrequirement, "Product"));

            // Open form and prepare to wrap content in "staySelection" container
            if (_.IF(_outer.isexternalbooking))
            {
                intbookingtype = _.VAL(_outer.booking_redirect);
            }
            else
            {
                intbookingtype = _.VAL(_outer.booking_local);
            }

            _.CALL(this, _outer, "RenderBookingInfoForm", _.ARGS.Ref(po, v2 => { po = v2; }).Ref(intprodkey, v3 => { intprodkey = v3; }).Ref(objrendersettings, v4 => { objrendersettings = v4; }).Ref(intbookingtype, v5 => { intbookingtype = v5; }).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null));

            _.CALL(this, po, "Write", _.ARGS.Val("<div class=\"staySelection\">"));

            // Try to pull requirement info from Request
            istaynum = (Int16)1;
            ithisreqmnt = (Int16)0;
            var enumerationContent = _.ENUMERABLE(_.CALL(this, _outer.request, "Form")).GetEnumerator();
            while (true)
            {
                if (!enumerationContent.MoveNext())
                    break;
                item = enumerationContent.Current;
                //## Loop through only units
                if (_.IF(_.EQ(_.NullableSTR(_.LEFT(item, (Int16)5)), "unit_")))
                {

                    strtemp = _.VAL(_.RIGHT(item, _.SUBT(_.LEN(item), (Int16)5)));
                    if (_.IF(_.GT(_.NullableNUM(_.INSTR(strtemp, "_")), (Int16)0)))
                    {
                        // Linked unit
                        iunitkey = _.CLNG(_.RIGHT(strtemp, _.SUBT(_.LEN(strtemp), _.INSTR(strtemp, "_"))));
                        ilinkedunitkey = _.CLNG(_.LEFT(strtemp, _.SUBT(_.INSTR(strtemp, "_"), (Int16)1)));
                    }
                    else
                    {
                        iunitkey = _.CLNG(strtemp);
                        ilinkedunitkey = (Int16)0;
                    }

                    iunitqty = _.CLNG(_.CONCAT("0", _.CALL(this, _outer.request, _.ARGS.Ref(item, v6 => { item = v6; }))));
                    iunitminoccupancy = _.CLNG(_.CONCAT("0", _.CALL(this, _outer.request, _.ARGS.Val(_.CONCAT("minoccu_", strtemp)))));
                    iunitmaxcapacity = _.CLNG(_.CONCAT("0", _.CALL(this, _outer.request, _.ARGS.Val(_.CONCAT("maxcap_", strtemp)))));

                    strunitname = _.VAL(_.CALL(this, _outer.request, _.ARGS.Val(_.CONCAT("name_", strtemp))));
                    stravailclassid = _.VAL(_.CALL(this, _outer.request, _.ARGS.Val(_.CONCAT("availclass_", strtemp))));
                    if (_.IF(_.GT(_.NullableNUM(iunitqty), (Int16)0)))
                    {
                        var loopEnd = _.NUM(iunitqty);
                        var loopStart = _.NUM((Int16)1, loopEnd);
                        if (_.StrictLTE(loopStart, loopEnd))
                        {
                            for (i = loopStart; _.StrictLTE(i, loopEnd); i = _.ADD(i, (Int16)1))
                            {
                                ithisreqmnt = _.ADD(ithisreqmnt, (Int16)1);
                                _.CALL(this, _outer, "BookingUI_RenderNewReq_AvailCal", _.ARGS.Ref(intbookingtype, v7 => { intbookingtype = v7; }).Ref(iunitkey, v8 => { iunitkey = v8; }).Ref(strunitname, v9 => { strunitname = v9; }).Ref(iunitminoccupancy, v10 => { iunitminoccupancy = v10; }).Ref(iunitmaxcapacity, v11 => { iunitmaxcapacity = v11; }).Ref(stravailclassid, v12 => { stravailclassid = v12; }).Ref(istaynum, v13 => { istaynum = v13; }).Ref(ithisreqmnt, v14 => { ithisreqmnt = v14; }).Ref(po, v15 => { po = v15; }));
                            }
                        }
                        if (_.IF(_.GT(_.NullableNUM(ilinkedunitkey), (Int16)0)))
                        {
                            _.CALL(this, po, "write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"linked_", iunitkey, "\"  value=\"", ilinkedunitkey, "\" />")));
                        }
                    }
                }
            }

            // If successfully received requirement data, complete form - otherwise render error
            if (_.IF(_.GT(_.NullableNUM(ithisreqmnt), (Int16)0)))
            {
                _.CALL(this, po, "write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"availcal\" value=\"", _.CALL(this, _outer.request, _.ARGS.Val("availcal")), "\" />")));
                _.CALL(this, po, "write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"_nStays\" value=\"", istaynum, "\" />")));
                _.CALL(this, po, "write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"_nReqs\" value=\"", ithisreqmnt, "\" />")));

                // Close pnStayReqmntRslts div
                _.CALL(this, po, "write", _.ARGS.Val("</div>"));

                _.CALL(this, _outer, "BookingUI_RenderButtons", _.ARGS.Ref(istaynum, v16 => { istaynum = v16; }).Ref(po, v17 => { po = v17; }).Val(false));

                // Close StayCandidateItem div
                _.CALL(this, po, "write", _.ARGS.Val("</div>"));
            }
            else
            {
                _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/availcalendar/nounitsselectederror").Val("<h2>Error</h2><p class=\"error\">No units selected. Please click on the back button to return to the previous page and select the units you wish to book.</p>"))));
            }

            // Close "staySelection" div and form
            _.CALL(this, po, "Write", _.ARGS.Val("</div>"));
            _.CALL(this, po, "Write", _.ARGS.Val("</form>"));
            if (_.IF(_.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("Booking_ChildPricing"))))
            {
                _.CALL(this, po, "Write", _.ARGS.Val("<script type=\"text/javascript\">"));
                _.CALL(this, po, "Write", _.ARGS.Val("NewMind.ETWP.Booking.UnitSelectionChildPricingGuests.Init();"));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("</", "script>")));
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
        public object bookingui_staymain(ref object objrendersettings, ref object objdata)
        {
            object BookingUI_StayMain_retVal = null;
            //most of these render functions rely on global variables, rather than trying to refactor them out for now ill create some globals
            //this needs refactoring

            // 2011-08-09 DWR: Expect the BookingRequirement in objRenderSettings to be read-only (since it usually comes from Page.Functions.GetSharedObject),
            // so replace it with an editable version (since some methods in here try to mess about with properties on it)
            _.SET(_.OBJ(_.CALL(this, _outer, "GetEditableBookingRequirement", _.ARGS.Val(_.CALL(this, objrendersettings, "BookingRequirement")))), this, objrendersettings, "BookingRequirement");

            _outer.isvbpollingenabled = _.VAL(_.CALL(this, objrendersettings, "IsVBPollingEnabled"));
            _outer.brenderascalendar = _.VAL(_.CALL(this, objrendersettings, "RenderAsCalendar"));
            if (_.IF(_.IS(objdata, VBScriptConstants.Nothing)))
            {
                // If couldn't retrieve product, report no availability - this will happen if the
                // availability criteria can (no longer) be met
                object byrefalias = objrendersettings;
                try
                {
                    _.CALL(this, _outer, "RenderNoAvailElement", _.ARGS.Ref(byrefalias, v18 => { byrefalias = v18; }));
                }
                finally { objrendersettings = byrefalias; }
                return BookingUI_StayMain_retVal;
            }

            if (_.IF(_.CALL(this, objrendersettings, "LegacyRender")))
            {
                // Acco or Ticketing w/out VB Polling Enabled: Results from single Supplier (either
                // local OR FrontDesk for Acco, only local applies for Tickets)
                object byrefalias2 = objdata, byrefalias3 = objrendersettings;
                try
                {
                    _.CALL(this, _outer, "BookingUI_StayMain_Legacy", _.ARGS.Ref(byrefalias2, v19 => { byrefalias2 = v19; }).Ref(byrefalias3, v20 => { byrefalias3 = v20; }));
                }
                finally { objdata = byrefalias2; objrendersettings = byrefalias3; }
            }
            else
            {
                // Acco w/ VB Polling Enabled: Results from multiple Suppliers
                // - Not supported when handling Conference Bookings, these are local only (but when
                //   an OfferKey is set, IsVBPollingEnabled is put to False - see PreRender)
                object byrefalias4 = objdata, byrefalias5 = objrendersettings;
                try
                {
                    _.CALL(this, _outer, "BookingUI_StayMain_Polling", _.ARGS.Ref(byrefalias4, v21 => { byrefalias4 = v21; }).Ref(byrefalias5, v22 => { byrefalias5 = v22; }));
                }
                finally { objdata = byrefalias4; objrendersettings = byrefalias5; }
            }
            return BookingUI_StayMain_retVal;
        }

        // ====================================================================================================
        // RENDER: Write out form with hidden input fields used for internal or FrontDesk bookings
        // - This will open the form, but the caller must close it
        // ====================================================================================================
        // Note: We need to pass intProdKey into here as we may not have an objProduct reference
        // (eg. if called by BookingUI_StayMain_AvailCal)
        public object renderbookinginfoform(object po, object intprodkey, object objrendersettings, object intbookingtype, object strsupplierid, object strsuppliername, object strsuppliereviivoname, object strsupplierdeeplinkquality, object strsupplierlogo, object inteviivosearchindustryclassification)
        {
            object RenderBookingInfoForm_retVal = null;
            object strposturl = null;
            object strformclass = null;
            object strnextstage = null;

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
            if (_.IF(_.EQ(intbookingtype, _outer.booking_redirect)))
            {
                strformclass = "FrmUnitOptionsExt";
            }
            else
            {
                strformclass = "FrmUnitOptions";
            }

            // What booking stage is next?
            // - If not external, go to checkout regardless of VB Polling setting.
            // - If IS external, branch off differently (VB Polling goes to a separate switcher stage, non-
            //   VB-Polling will redirect to the other site).
            // While we're here, retrieve POST url (secure for checkout, standard otherwise)
            if (_.IF(_.OR(_.EQ(intbookingtype, _outer.booking_local), _.EQ(intbookingtype, _outer.booking_eviivo))))
            {
                strnextstage = "checkout";
                strposturl = _.CONCAT(_.CALL(this, _outer, "GetPostUrl", _.ARGS.Val(true)), "/", strnextstage);
            }
            else if (_.IF(_.EQ(intbookingtype, _outer.booking_redirect)))
            {
                //strNextStage = "redirect"
                strnextstage = "checkout";
                //This should stay as "checkout" until 1.4 is updated to recognise "redirect" stage
                strposturl = _.VAL(_.CALL(this, _outer.page, "PageInfo", "GetUrlFromPageID", _.ARGS.Val("EXTBOOKPROMPT")));
                if (_.IF(_.ISNULL(strposturl)))
                {
                    _.CALL(this, _outer.page, "PrintTraceWarning", _.ARGS.Val("RenderBookingInfoForm: Unable to locate page EXTBOOKPROMPT, default to current page - is this correct behaviour??"));
                    strposturl = _.VAL(_.CALL(this, _outer.page, "URL", "Real"));
                }
            }
            else if (_.IF(_.EQ(intbookingtype, _outer.booking_pollingredirect)))
            {
                // 2014-06-19 DWR: We have historically used the SupplierEviivoName for the URL segment, although it used to be labelled strSupplierName since
                // the values were getting set incorrectly. SupplierEviivoName seems like the most appropriate option since it will be a text-friendly string
                // value and so not have dots or spaces or whatever (and so be good for use in a URL).
                strnextstage = "pollingexit";
                strposturl = _.CONCAT(_.CALL(this, _outer, "GetPostUrl", _.ARGS.Val(false)), "/pollingexit/", strsuppliereviivoname);
            }
            else
            {
                _.RAISEERROR(VBScriptConstants.vbObjectError, "ETWP.BookingUnitSelection", _.CONCAT("RenderBookingInfoForm: Invalid intBookingType value (", intbookingtype, ")"));
            }

            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<form action=\"", strposturl, "\" ")));
            if (_.IF(_.AND(_.NOT(_outer.isvbpollingenabled), _.EQ(_.NullableNUM(_.CALL(this, objrendersettings, "BookingRequirement", "FlexibleRange")), (Int16)0))))
            {
                // Can't have ids when VB Polling enabled as we might be rendering out multiple of these forms.
                // 2008-11-10 DWR: This is similarly the case for fuzzy searching. I don't we have any working
                // Enterprise fuzzy-searching sites, so don't need to worry about breaking styling by removing
                // this id in this case.
                _.CALL(this, po, "Write", _.ARGS.Val("id=\"FrmUnitOptions\" "));
            }

            //#MJ's Reasoning -	In order for us to jump to unit selection in a tab it must have a name, however only the first form should have this
            if (_.IF(_.NOT(_outer.bformrendered)))
            {
                _.CALL(this, po, "Write", _.ARGS.Val("name=\"FrmUnitOptions\" "));
                _outer.bformrendered = true;
            }
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("class=\"", strformclass, "\" method=\"post\">")));

            // Open container around common form values
            _.CALL(this, po, "Write", _.ARGS.Val("<div>"));

            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"stage\" value=\"", strnextstage, "\" />")));

            // Need to override market source if viewing site via widget
            if (_.IF(_.CALL(this, _outer.page, "WidgetView")))
            {
                if (_.IF(_.EQ(intbookingtype, _outer.booking_redirect)))
                {
                    // External bookings visit a preliminary redirect page first, which we want to be decluttered when in a widget
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"widget_marketsource\" value=\"", _.CALL(this, _outer.page, "WidgetMarketSource"), "\" />")));
                }
                else
                {
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"msource\" value=\"", _.CALL(this, _outer.page, "WidgetMarketSource"), "\" />")));
                }
                //this hidden field is to tell the checkout that weve come from a widget, and not a failed checkout validation
                _.CALL(this, po, "Write", _.ARGS.Val("<input type=\"hidden\" name=\"widget\" value=\"1\" />"));
            }

            // None of this applies to VB Polling, even if it IS an external booking - we go to an
            // interim stage before leaving the site
            if (_.IF(_.EQ(intbookingtype, _outer.booking_redirect)))
            {
                // NB: In "Conference Booking" mode (where OfferKey <> 0), we need to set the "channel" and "msource"
                //     values to different values (for msource, if there is no "ConfBookingMarketSourceID" set, it will
                //     fall back to using the site's main "MarketSourceID" source)
                _.CALL(this, po, "Write", _.ARGS.Val("<input type=\"hidden\" name=\"checkoutstage\" value=\"1\" />"));
                if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, objrendersettings, "BookingRequirement", "Offer")), (Int16)0)))
                {
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"channel\" value=\"", _.CALL(this, objrendersettings, "Channel"), "\" />")));
                }
                else
                {
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"channel\" value=\"", _.CALL(this, objrendersettings, "ConfBookingChannel"), "\" />")));
                }
                if (_.IF(_.NOT(_.CALL(this, _outer.page, "WidgetView"))))
                {
                    //Neeed to set market source override if redirecting to external site unless set above due to widgetview
                    if (_.IF(_.OR(_.EQ(_.NullableNUM(_.CALL(this, objrendersettings, "BookingRequirement", "Offer")), (Int16)0), _.EQ(_.NullableSTR(_.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("ConfBookingMarketSourceID"))), ""))))
                    {
                        _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"msource\" value=\"", _.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("MarketSourceID")), "\" />")));
                    }
                    else
                    {
                        _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"msource\" value=\"", _.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("ConfBookingMarketSourceID")), "\" />")));
                    }
                }
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"bookchannel\" value=\"", _.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("Booking_ChannelID")), "\" />")));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"reposturl\" value=\"", _outer.strextbookurl, "\" />")));
                // 2009-09-21 DWR: New field to pass in so that the receiving site recognises booking as having
                // come from another site (so it can update appropriate Provider Stats)
                _.CALL(this, po, "Write", _.ARGS.Val("<input type=\"hidden\" name=\"ForcedExternalBooking\" value=\"1\" />"));
            }

            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"product\" value=\"", intprodkey, "\" />")));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"isostartdate\" value=\"", _.CALL(this, _outer.page, "Functions", "Dates", "ISODate", _.ARGS.Val(_.CALL(this, objrendersettings, "BookingRequirement", "VisitDate"))), "\" />")));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"nights\" value=\"", _.CALL(this, objrendersettings, "BookingRequirement", "Nights"), "\" />")));

            // We need all this when using VB Polling, even it it is an external booking, as we aren't
            // going to leave the site yet (there's an interim stage)
            if (_.IF(_.NOTEQ(intbookingtype, _outer.booking_redirect)))
            {
                // NB: "package" parameter removed - it's now passed as "offer", and only when
                // customer is going for a "Conference Booking" discount product.
                // 2008-11-07 DWR: This used to referer to a "strRewriteUrl" value that was never defined.
                // So we'll pass in blank. Pretty sure it's not used anyway.
                _.CALL(this, po, "Write", _.ARGS.Val("<input type=\"hidden\" name=\"preUrl\" value=\"\" />"));
                // 2008-11-07 DWR: If we've got non-precise results from a fuzzy search, we'll render this
                // form out and use the actual StartDate / NumNights combination that the fuzzy results
                // offered. So we just pass these to the checkout stage, and set "fuzzy" to zero.
                _.CALL(this, po, "Write", _.ARGS.Val("<input type=\"hidden\" name=\"fuzzy\" value=\"0\" />"));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"lng\" value=\"", _.CALL(this, _outer.page, "Language", "LanguageCultureKey"), "\" />")));

                // NB: OfferKey is required for products in the "Conference Booking" functionality as
                // it lets the checkout object know that we should be looking for the product on the
                // "Conference Booking Channel" instead of the standard "website" channel. If this
                // ever needed to work with the ExternalBooking, we would need to pass out the
                // conference channel in the IsExternalBooking section above, but since this is
                // only being supported by the internal Newmind booking, it's not an issue.
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"offer\" value=\"", _.CALL(this, objrendersettings, "BookingRequirement", "Offer"), "\" />")));

                // Pass in the current convert-to-currency value (this will have been held in the session
                // up to this point, but we may be about to leave the site when this form is posted, so
                // will need to send the value as a hidden input instead of relying on session)
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"CurrencyConvertTo\" value=\"", _.CALL(this, _outer.page, "Functions", "Money", "GetCurrencyCodeOverride", _.ARGS.Val(_.CALL(this, _outer.page, "Site", "LCCurrencyKey"))), "\" />")));
            }

            // If we're dealing with a VB Polling External Supplier, write out the Supplier id, name and
            // deep-link-quality as well (this is the number of rooms that the supplier can handle in
            // deep-linking situations)
            if (_.IF(_.EQ(intbookingtype, _outer.booking_pollingredirect)))
            {
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"SupplierId\" value=\"", strsupplierid, "\" />")));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"SupplierName\" value=\"", strsuppliername, "\" />")));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"SupplierLogo\" value=\"", strsupplierlogo, "\" />")));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"SupplierEviivoName\" value=\"", strsuppliereviivoname, "\" />")));

                _.CALL(this, po, "Write", _.ARGS.Val("<input type=\"hidden\" name=\"EviivoSearchIndustryClassification\" value=\""));
                if (_.IF(_.ISNUMERIC(inteviivosearchindustryclassification)))
                {
                    _.CALL(this, po, "Write", _.ARGS.Ref(inteviivosearchindustryclassification, v23 => { inteviivosearchindustryclassification = v23; }));
                }
                else
                {
                    _.CALL(this, po, "Write", _.ARGS.Val("0"));
                }
                _.CALL(this, po, "Write", _.ARGS.Val("\" />"));

                if (_.IF(_.ISNULL(strsupplierdeeplinkquality)))
                {
                    strsupplierdeeplinkquality = "";
                }
                else
                {
                    strsupplierdeeplinkquality = _.VAL(_.TRIM(strsupplierdeeplinkquality));
                }
                if (_.IF(_.NOT(_.ISNUMERIC(strsupplierdeeplinkquality))))
                {
                    strsupplierdeeplinkquality = "-1";
                }
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"SupplierDeepLinkQuality\" value=\"", strsupplierdeeplinkquality, "\" />")));
            }

            // Append in the "Nominal Units" from Request collection or objUnitReqDictFromBookUrl (ie. "roomReq_1", "roomReq_2", etc..)
            //#MJ TODO need to call the new function
            _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer, "GenerateRequirementFormData", _.ARGS.Val(_.CALL(this, objrendersettings, "BookingRequirement")))));
            // Close common form value container
            _.CALL(this, po, "Write", _.ARGS.Val("</div>"));

            return RenderBookingInfoForm_retVal;
        }

        //generates a string of room requirement details in a format suitable for use in a form i.e hidden inputs ;)
        public object generaterequirementformdata(ref object objaccosearchrequirement)
        {
            object GenerateRequirementFormData_retVal = null;
            object dictkeyvalues = null;
            object aryformatteddata = null;
            object i = null;
            object key = null;
            //get our key value data dictionary
            object byrefalias6 = objaccosearchrequirement;
            try
            {
                dictkeyvalues = _.OBJ(_.CALL(this, _outer.page, "Functions", "Booking", "GenerateRequirementKeyValueData", _.ARGS.Ref(byrefalias6, v24 => { byrefalias6 = v24; })));
            }
            finally { objaccosearchrequirement = byrefalias6; }
            //create an array to hold our formatted data in which is the same size of the dictionary
            aryformatteddata = _.NEWARRAY(new object[] { _.SUBT(_.CALL(this, dictkeyvalues, "Count"), (Int16)1) });
            //spin through our output array and add the formatted items in the format {key}={value}
            i = (Int16)0;
            var enumerationContent2 = _.ENUMERABLE(_.CALL(this, dictkeyvalues, "Keys")).GetEnumerator();
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
                    _.SET(_.CONCAT("<input type=\"hidden\" name=\"", key, "\" value=\"", _.CALL(this, dictkeyvalues, "Item", _.ARGS.Ref(key, v27 => { key = v27; })), "\" />", VBScriptConstants.vbCrLf), this, aryformatteddata, null, _.ARGS.Ref(i, v26 => { i = v26; }));
                }
                i = _.ADD(i, (Int16)1);
            }
            //return our array as a string using an & as the joining character
            GenerateRequirementFormData_retVal = _.JOIN(aryformatteddata);
            return GenerateRequirementFormData_retVal;
        }

        public object getposturl(object bsecure)
        {
            object GetPostUrl_retVal = null;
            object strposturl = null;
            object strurl = null;

            if (_.IF(bsecure))
            {
                strposturl = _.VAL(_.CALL(this, _outer.page, "Site", "SecureHostName"));
            }
            else
            {
                strposturl = _.VAL(_.CALL(this, _outer.page, "URL", "FullHostName"));
            }
            while (_.IF(_.EQ(_.NullableSTR(_.RIGHT(strposturl, (Int16)1)), "/")))
            {
                strposturl = _.VAL(_.LEFT(strposturl, _.SUBT(_.LEN(strposturl), (Int16)1)));
            }

            strurl = _.VAL(_.CALL(this, _outer.page, "PageInfo", "GetUrlFromPageID", _.ARGS.Val("BOOKONLINE")));
            if (_.IF(_.ISNULL(strurl)))
            {
                _.CALL(this, _outer.page, "PrintTraceWarning", _.ARGS.Val("GetPostUrl: Unable to locate page BOOKONLINE, default to current page - is this correct behaviour??"));
                strurl = _.VAL(_.CALL(this, _outer.page, "URL", "Real"));
            }

            if (_.IF(_.NOTEQ(_.NullableSTR(_.LEFT(strurl, (Int16)1)), "/")))
            {
                strurl = _.CONCAT("/", strurl);
            }

            strposturl = _.CONCAT(strposturl, strurl);
            if (_.IF(_.AND(_.NOTEQ(_.NullableSTR(_.UCASE(_.LEFT(strposturl, (Int16)7))), "HTTP://"), _.NOTEQ(_.NullableSTR(_.UCASE(_.LEFT(strposturl, (Int16)8))), "HTTPS://"))))
            {
                strposturl = _.CONCAT("http://", strposturl);
            }

            GetPostUrl_retVal = _.VAL(strposturl);
            return GetPostUrl_retVal;
        }

        // SUMMARY: render new requirement UI from avail calendar
        // [ireqSz]: ADO unit recordset from availability object
        // [aiStayNum]: integer stay index
        // [aiThisReqmnt]: integer requirement number (from recordset)
        public object bookingui_rendernewreq_availcal(ref object intbookingtype, ref object iunitkey, ref object strunitname, ref object iunitminoccupancy, ref object iunitmaxcapacity, ref object asavailclassid, ref object aistaynum, ref object aithisreqmnt, ref object po)
        {
            object BookingUI_RenderNewReq_AvailCal_retVal = null;
            object iguest = null;
            object strguestsfor = null;
            object stradultstitle = null;
            object stradults = null;
            object strchildrentitle = null;
            object strchildren = null;
            object strguestsand = null;
            object icount = null;
            object agevalue = null;
            object strgueststitle = null;
            object strguests = null;
            // on first ever call [aiThisReqmnt]=1, on subsequent calls we must close previous [pnStayReqmnt] and [pnStayReqmntRslts] DIVs

            if (_.IF(_.GT(_.NullableNUM(aithisreqmnt), (Int16)1)))
            {
                _.CALL(this, po, "Write", _.ARGS.Val("</div></div>"));
            }

            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<div class=\"pnStayReqmnt\">", VBScriptConstants.vbCrLf)));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<div class=\"pnStayReqmntTtl\">", VBScriptConstants.vbCrLf)));
            object byrefalias7 = asavailclassid;
            try
            {
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<div Class=\"pnStayReqmntRoom\">Room ", aithisreqmnt, " - ", strunitname, _.CALL(this, _outer, "BookingUI_AvailClassIcon", _.ARGS.Ref(byrefalias7, v29 => { byrefalias7 = v29; })), " <br/></div>")));
            }
            finally { asavailclassid = byrefalias7; }

            if (_.IF(_.OR(_.EQ(_.NullableNUM(iunitminoccupancy), (Int16)0), _.EQ(_.NullableSTR(iunitminoccupancy), ""))))
            {
                iunitminoccupancy = (Int16)1;
            }

            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<div Class=\"pnStayReqmntGuests\">", VBScriptConstants.vbCrLf)));
            if (_.IF(_.EQ(iunitmaxcapacity, iunitminoccupancy)))
            {
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("For ", iunitmaxcapacity, " guests <input type=\"hidden\" name=\"roomReq_", aithisreqmnt, "\" value=\"", iunitmaxcapacity, "\"/>")));
            }
            else
            {
                strguestsfor = _.VAL(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/guestrequirement/for").Val("for")));
                //alas child pricing is different
                if (_.IF(_.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("Booking_ChildPricing"))))
                {

                    stradultstitle = _.VAL(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/guestrequirement/adults/selecttitle").Val("Please specify the number of adults in this room.")));
                    stradults = _.VAL(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/guestrequirement/adults/adult(s)").Val("adult(s)")));

                    _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT(strguestsfor, " <select class=\"adults\" name=\"roomReq_", aithisreqmnt, "_adults\" title=\"", stradultstitle, "\"> ")));
                    var loopEnd2 = _.NUM(iunitmaxcapacity);
                    var loopStart2 = _.NUM(iunitminoccupancy, loopEnd2, (Int16)1);
                    if (_.StrictLTE(loopStart2, loopEnd2))
                    {
                        for (iguest = loopStart2; _.StrictLTE(iguest, loopEnd2); iguest = _.ADD(iguest, (Int16)1))
                        {
                            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<option value=\"", iguest, "\">", iguest, "</option> ")));
                        }
                    }
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("</select> ", stradults)));

                    strchildrentitle = _.VAL(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/guestrequirement/children/selecttitle").Val("Please specify the number of children in this room.")));
                    strchildren = _.VAL(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/guestrequirement/children/children").Val("children")));
                    strguestsand = _.VAL(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("and").Val("and")));
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT(" ", strguestsand, " <select class=\"children\" name=\"roomReq_", aithisreqmnt, "_children\" title=\"", strchildrentitle, "\"> ")));

                    var loopEnd3 = _.NUM(_.SUBT(iunitmaxcapacity, (Int16)1));
                    var loopStart3 = _.NUM((Int16)0, loopEnd3, (Int16)1);
                    if (_.StrictLTE(loopStart3, loopEnd3))
                    {
                        for (iguest = loopStart3; _.StrictLTE(iguest, loopEnd3); iguest = _.ADD(iguest, (Int16)1))
                        {
                            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<option value=\"", iguest, "\">", iguest, "</option> ")));
                        }
                    }
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("</select> ", strchildren)));

                    _.CALL(this, po, "WriteLine", _.ARGS.Val("<span class=\"label childrenageslabel\">Child Ages</span>"));
                    _.CALL(this, po, "WriteLine", _.ARGS.Val("<span class=\"field childrenagesfield\">"));

                    var loopEnd4 = _.NUM(_.SUBT(iunitmaxcapacity, (Int16)1));
                    var loopStart4 = _.NUM((Int16)0, loopEnd4, (Int16)1);
                    if (_.StrictLTE(loopStart4, loopEnd4))
                    {
                        for (icount = loopStart4; _.StrictLTE(icount, loopEnd4); icount = _.ADD(icount, (Int16)1))
                        {
                            _.CALL(this, po, "WriteLine", _.ARGS.Val("<span class=\"childageWrapper\">"));
                            _.CALL(this, po, "WriteLine", _.ARGS.Val(_.CONCAT(VBScriptConstants.vbTab, "<span class=\"label childagelabel\">Child Age ", _.ADD(icount, (Int16)1), "</span>")));
                            _.CALL(this, po, "WriteLine", _.ARGS.Val(_.CONCAT(VBScriptConstants.vbTab, "<span class=\"field childagefield\">")));
                            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<select class=\"\" name=\"roomReq_", aithisreqmnt, "_children_childage", icount, "\">")));
                            for (iguest = (Int16)0; _.StrictLTE(iguest, 18); iguest = _.ADD(iguest, (Int16)1))
                            {
                                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<option value=\"", iguest, "\">", iguest, "</option> ")));
                            }
                            _.CALL(this, po, "Write", _.ARGS.Val("</select> "));
                            _.CALL(this, po, "WriteLine", _.ARGS.Val(_.CONCAT(VBScriptConstants.vbTab, "</span>")));
                            _.CALL(this, po, "WriteLine", _.ARGS.Val("</span>"));
                        }
                    }
                    _.CALL(this, po, "WriteLine", _.ARGS.Val("</span>"));
                }
                else
                {
                    strgueststitle = _.VAL(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/guestrequirement/selecttitle").Val("Please specify the number of guests in this room.")));
                    strguests = _.VAL(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/guestrequirement/guest(s)").Val("guest(s)")));

                    _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT(strguestsfor, " <select name=\"roomReq_", aithisreqmnt, "\" title=\"", strgueststitle, "\"> ")));
                    var loopEnd5 = _.NUM(iunitmaxcapacity);
                    var loopStart5 = _.NUM(iunitminoccupancy, loopEnd5, (Int16)1);
                    if (_.StrictLTE(loopStart5, loopEnd5))
                    {
                        for (iguest = loopStart5; _.StrictLTE(iguest, loopEnd5); iguest = _.ADD(iguest, (Int16)1))
                        {
                            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<option value=\"", iguest, "\">", iguest, "</option> ")));
                        }
                    }
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("</select> ", strguests)));
                }
            }
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("</div>", VBScriptConstants.vbCrLf)));

            if (_.IF(_.EQ(_.NullableSTR(intbookingtype), "ticketing")))
            {
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"unit_", iunitkey, "\"  value=\"", aithisreqmnt, "\" />")));
            }
            else
            {
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"unit_", aistaynum, "_", aithisreqmnt, "\"  value=\"", iunitkey, "\" />")));
            }
            return BookingUI_RenderNewReq_AvailCal_retVal;
        }

        // SUMMARY: Draw availability month calendar
        // [sbCalendars]:  ASP [nmStringBuilder] object instance output string
        // [dCalStartDflt]: date default calendar start date
        // <retval>: string month available stays details JSON data
        public object bookingui_renderavailcal(ref object sbcalendars, ref object objdictavaistays, ref object bstarted)
        {
            object BookingUI_RenderAvailCal_retVal = null;
            object strclassmonth = null;
            object dstart1 = null;
            object aryavailstayskeys = null;

            strclassmonth = "MonthWrapper";
            if (_.IF(_.NOT(bstarted)))
            {
                _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val("<div class=\"CalendarsWrapper\">"));
                _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val(_.CONCAT("<div class=\"instruction\">", _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/availcalendar/instruction").Val("Please select an available stay from the calendars below. Clicking on a highlighted start day for a stay will show the stay details such as the units available, price, etc.")), "</div>")));
                strclassmonth = _.CONCAT(strclassmonth, " currentmonth");
            }
            else
            {
                strclassmonth = _.CONCAT(strclassmonth, " nextmonth");
            }

            if (_.IF(_.GT(_.NullableNUM(_.CALL(this, objdictavaistays, "Count")), (Int16)0)))
            {
                aryavailstayskeys = _.VAL(_.CALL(this, objdictavaistays, "Keys"));
                dstart1 = _.REPLACE(_.CALL(this, aryavailstayskeys, _.ARGS.Val((Int16)0)), "sd_", "");
                _.ERASE(aryavailstayskeys, v31 => { aryavailstayskeys = v31; });
            }
            else
            {
                dstart1 = _.DATE();
            }

            object byrefalias8 = sbcalendars, byrefalias9 = objdictavaistays;
            try
            {
                _.CALL(this, _outer, "BookingUI_RenderCalendarMonthWithAvailability", _.ARGS.Ref(byrefalias8, v32 => { byrefalias8 = v32; }).Ref(dstart1, v33 => { dstart1 = v33; }).Ref(strclassmonth, v34 => { strclassmonth = v34; }).Ref(byrefalias9, v35 => { byrefalias9 = v35; }));
            }
            finally { sbcalendars = byrefalias8; objdictavaistays = byrefalias9; }

            // using a global count so we can track how many calendars have been added to the stringbuilder for the prev/next buttons
            // doing this now because of the recursive nature of this function
            _outer.g_inumberofcalendarsrendered = _.ADD(_outer.g_inumberofcalendarsrendered, (Int16)1);

            //Check if we have stays left and render then as another calendar
            if (_.IF(_.GT(_.NullableNUM(_.CALL(this, objdictavaistays, "Count")), (Int16)0)))
            {
                object byrefalias10 = sbcalendars, byrefalias11 = objdictavaistays;
                try
                {
                    _.CALL(this, _outer, "BookingUI_RenderAvailCal", _.ARGS.Ref(byrefalias10, v36 => { byrefalias10 = v36; }).Ref(byrefalias11, v37 => { byrefalias11 = v37; }).Val(true));
                }
                finally { sbcalendars = byrefalias10; objdictavaistays = byrefalias11; }
            }
            else
            {
                //not sure if this should be dStart1 - was dStart
                object byrefalias12 = sbcalendars;
                try
                {
                    _.CALL(this, _outer, "BookingUI_RenderAvailCalLinks", _.ARGS.Ref(dstart1, v38 => { dstart1 = v38; }).Ref(byrefalias12, v39 => { byrefalias12 = v39; }));
                }
                finally { sbcalendars = byrefalias12; }
                object byrefalias13 = sbcalendars;
                try
                {
                    _.CALL(this, _outer, "BookingUI_RenderAvailCalKey", _.ARGS.Ref(byrefalias13, v40 => { byrefalias13 = v40; }));
                }
                finally { sbcalendars = byrefalias13; }
                _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val("</div>"));

            }

            return BookingUI_RenderAvailCal_retVal;
        }

        public object bookingui_rendercalendarmonth(ref object sbcalendars, object dfirstdayofmonth, object strwrapperclass)
        {
            object BookingUI_RenderCalendarMonth_retVal = null;
            object byrefalias14 = sbcalendars;
            try
            {
                _.CALL(this, _outer, "BookingUI_RenderCalendarMonthWithAvailability", _.ARGS.Ref(byrefalias14, v41 => { byrefalias14 = v41; }).Ref(dfirstdayofmonth, v42 => { dfirstdayofmonth = v42; }).Ref(strwrapperclass, v43 => { strwrapperclass = v43; }).Val(VBScriptConstants.Nothing));
            }
            finally { sbcalendars = byrefalias14; }
            return BookingUI_RenderCalendarMonth_retVal;
        }

        public object bookingui_rendercalendarmonthwithavailability(ref object sbcalendars, object dfirstdayofmonth, object strwrapperclass, object objdictavailstays)
        {
            object BookingUI_RenderCalendarMonthWithAvailability_retVal = null;
            object iweekstartday = null;
            object iweekdaycalstart = null;
            object iweekdaycalend = null;
            object dcalstart = null;
            object dcalend = null;
            object strthismonthyear = null;
            object strtablesummary = null;
            object strheadercellclass = null;
            object i = null;
            object icellcount = null;
            object bfirstcell = null;
            object blastcell = null;
            object ddate = null;
            object bstartnewstay = null;
            object bstayindicative = null;
            object strstaynumber = null;
            object iday = null;
            object iprepadding = null;
            object j = null;
            object strdisplaytext = null;
            object strdaycellclass = null;
            object arystay = null;
            object stravailtype = null;
            object strindicativeicon = null;
            object ipostpadding = null;
            object k = null;

            iweekstartday = (Int16)1; //Monday
            iweekdaycalstart = _.MOD(_.ADD(iweekstartday, (Int16)1), (Int16)7);
            iweekdaycalend = _.MOD(iweekstartday, (Int16)7);

            dcalstart = _.VAL(_.CALL(this, _outer.page, "Functions", "Dates", "fn_GetFirstDateOfMonth", _.ARGS.Ref(dfirstdayofmonth, v44 => { dfirstdayofmonth = v44; })));
            dcalend = _.VAL(_.CALL(this, _outer.page, "Functions", "Dates", "fn_GetLastDateOfMonth", _.ARGS.Ref(dfirstdayofmonth, v45 => { dfirstdayofmonth = v45; })));
            strthismonthyear = _.CONCAT(_.CALL(this, _outer.page, "Functions", "Dates", "GetMonthNameAbbr", _.ARGS.Val(_.MONTH(dcalstart))), " ", _.YEAR(dcalstart));
            strtablesummary = _.CONCAT(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/availcalendar/availabilitycalendarfor").Val("Availability calendar for")), " ", strthismonthyear);

            _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val(_.CONCAT("<div id=\"Cal_", _.CALL(this, _outer.page, "Functions", "Dates", "ISODate", _.ARGS.Ref(dcalstart, v46 => { dcalstart = v46; })), "\" class=\"", strwrapperclass, "\">")));
            _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val(_.CONCAT("<table id=\"Tbl_", _.CALL(this, _outer.page, "Functions", "Dates", "ISODate", _.ARGS.Ref(dcalstart, v48 => { dcalstart = v48; })), "\" class=\"availabilityCalendar\" summary=\"", strtablesummary, "\" >")));
            _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val("<thead>"));
            _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val("<tr>"));
            _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val(_.CONCAT("<th colspan=\"8\">", strthismonthyear, "</th>")));
            _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val("</tr>"));

            strheadercellclass = "";
            var loopEnd6 = _.NUM(_.ADD(iweekstartday, (Int16)6));
            var loopStart6 = _.NUM(iweekstartday, loopEnd6, (Int16)1);
            if (_.StrictLTE(loopStart6, loopEnd6))
            {
                for (i = loopStart6; _.StrictLTE(i, loopEnd6); i = _.ADD(i, (Int16)1))
                {
                    if (_.IF(_.OR(_.EQ(_.NullableNUM(_.MOD(i, (Int16)7)), (Int16)6), _.EQ(_.NullableNUM(_.MOD(i, (Int16)7)), (Int16)0))))
                    {
                        strheadercellclass = " class=\"we\"";
                    }
                    _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val(_.CONCAT("<th", strheadercellclass, ">", _.CALL(this, _outer.page, "Functions", "Dates", "GetDayNameAbbr", _.ARGS.Val(_.WEEKDAY(_.MOD(_.ADD(i, (Int16)1), (Int16)7)))), "</th>")));
                }
            }

            _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val("</tr>"));
            _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val("</thead>"));
            _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val("<tbody>"));

            icellcount = (Int16)0;
            bfirstcell = true;
            blastcell = false;

            ddate = _.VAL(dcalstart);

            var loopEnd7 = _.NUM(_.DAY(dcalend));
            var loopStart7 = _.NUM(_.DAY(dcalstart), loopEnd7, (Int16)1);
            if (_.StrictLTE(loopStart7, loopEnd7))
            {
                for (iday = loopStart7; _.StrictLTE(iday, loopEnd7); iday = _.ADD(iday, (Int16)1))
                {
                    bstartnewstay = false;

                    if (_.IF(bfirstcell))
                    {
                        iprepadding = _.VAL(_.DATEDIFF("d", _.CALL(this, _outer.page, "Functions", "Dates", "fn_GetFirstDateOfWeek", _.ARGS.Val(_.CALL(this, _outer.page, "Functions", "Dates", "fn_GetFirstDateOfMonth", _.ARGS.Ref(dcalstart, v50 => { dcalstart = v50; }))).Ref(iweekdaycalstart, v51 => { iweekdaycalstart = v51; })), dcalstart));
                        if (_.IF(_.GT(_.NullableNUM(iprepadding), (Int16)0)))
                        {
                            var loopEnd8 = _.NUM(iprepadding);
                            var loopStart8 = _.NUM((Int16)1, loopEnd8);
                            if (_.StrictLTE(loopStart8, loopEnd8))
                            {
                                for (j = loopStart8; _.StrictLTE(j, loopEnd8); j = _.ADD(j, (Int16)1))
                                {
                                    _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val("<td></td>"));
                                    icellcount = _.ADD(icellcount, (Int16)1);
                                }
                            }
                        }
                        bfirstcell = false;
                    }

                    strdisplaytext = _.CONCAT("", iday);
                    strdaycellclass = "n";

                    if (_.IF(_.NOT(_.IS(objdictavailstays, VBScriptConstants.Nothing))))
                    {
                        if (_.IF(_.CALL(this, objdictavailstays, "Exists", _.ARGS.Val(_.CONCAT("sd_", ddate)))))
                        {
                            bstartnewstay = true;
                            //we expect value in the format [stayNo]_[indicative]
                            arystay = _.SPLIT(_.CALL(this, objdictavailstays, _.ARGS.Val(_.CONCAT("sd_", ddate))), "_");
                            strstaynumber = _.VAL(_.CALL(this, arystay, _.ARGS.Val((Int16)0)));
                            bstayindicative = _.CBOOL(_.CALL(this, arystay, _.ARGS.Val((Int16)1)));
                            _.CALL(this, objdictavailstays, "Remove", _.ARGS.Val(_.CONCAT("sd_", ddate)));
                            _.ERASE(arystay, v52 => { arystay = v52; });
                        }
                    }

                    if (_.IF(_.LT(ddate, _.DATE()))) //date is in the past
                    {

                        strdaycellclass = "p";

                    }
                    else if (_.IF(bstartnewstay))
                    {

                        if (_.IF(_.NOT(_.IS(objdictavailstays, VBScriptConstants.Nothing))))
                        {
                            stravailtype = "";
                            strindicativeicon = "";

                            if (_.IF(bstayindicative))
                            {
                                strdaycellclass = "i";
                                stravailtype = _.VAL(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/unconfirmedavailability").Val("Unconfirmed Availability")));
                                strindicativeicon = _.CONCAT("<img src=\"", _.CALL(this, _outer.page, "ImageResource", _.ARGS.Val("bookonline/icons/indicative").Val("/images/icon_indicative.gif")), "\" alt=\"", stravailtype, "\" class=\"icon\"/>");
                            }
                            else
                            {
                                strdaycellclass = "a";
                                stravailtype = _.VAL(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/confirmedavailability").Val("Confirmed Availability")));
                                strindicativeicon = _.CONCAT("<img src=\"", _.CALL(this, _outer.page, "ImageResource", _.ARGS.Val("bookonline/icons/allocated").Val("/images/icon_allocated.gif")), "\" alt=\"", stravailtype, "\" class=\"icon\"/>");
                            }

                            strdisplaytext = _.CONCAT("<a href=\"#stay_", strstaynumber, "\" class=\"calavailstay\" id=\"stay_", strstaynumber, "\">", _.DAY(ddate), "</a>", strindicativeicon);

                        }

                    }

                    if (_.IF(_.OR(_.EQ(_.NullableNUM(_.WEEKDAY(ddate)), (Int16)1), _.EQ(_.NullableNUM(_.WEEKDAY(ddate)), (Int16)7))))
                    {
                        strdaycellclass = _.CONCAT(strdaycellclass, " we");
                    }

                    _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val(_.CONCAT("<td class=\"", strdaycellclass, "\"><div>", strdisplaytext, "</div></td>")));

                    icellcount = _.ADD(icellcount, (Int16)1);

                    if (_.IF(_.EQ(ddate, dcalend)))
                    {
                        blastcell = true;
                    }

                    // This is for when the last day of the month is not the last day of the week and empty cells are put in place to fill the calendar days
                    if (_.IF(blastcell))
                    {
                        ipostpadding = _.VAL(_.DATEDIFF("d", dcalend, _.CALL(this, _outer.page, "Functions", "Dates", "fn_GetLastDateOfWeek", _.ARGS.Ref(dcalend, v53 => { dcalend = v53; }).Ref(iweekdaycalend, v54 => { iweekdaycalend = v54; }))));
                        if (_.IF(_.AND(_.GT(_.NullableNUM(ipostpadding), (Int16)0), _.LT(_.NullableNUM(ipostpadding), (Int16)7))))
                        {
                            var loopEnd9 = _.NUM(ipostpadding);
                            var loopStart9 = _.NUM((Int16)1, loopEnd9);
                            if (_.StrictLTE(loopStart9, loopEnd9))
                            {
                                for (k = loopStart9; _.StrictLTE(k, loopEnd9); k = _.ADD(k, (Int16)1))
                                {
                                    _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val("<td></td>"));
                                    icellcount = _.ADD(icellcount, (Int16)1);
                                }
                            }
                        }
                        blastcell = false;
                        bfirstcell = true;
                    }

                    if (_.IF(_.EQ(_.NullableNUM(_.MOD(icellcount, (Int16)7)), (Int16)0)))
                    {
                        _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val("</tr>"));
                    }

                    ddate = _.VAL(_.DATEADD("d", 1, ddate));
                }
            }

            _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val("</tbody>"));
            _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val("</table>"));
            _.CALL(this, sbcalendars, "AppendLine", _.ARGS.Val("</div>"));

            return BookingUI_RenderCalendarMonthWithAvailability_retVal;
        }

        public object bookingui_renderavailcalkey(ref object sb)
        {
            object BookingUI_RenderAvailCalKey_retVal = null;
            object strcalkey = null;
            strcalkey = _.VAL(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/availcalendar/calkey").Val("")));
            if (_.IF(_.NOTEQ(_.NullableSTR(_.TRIM(strcalkey)), "")))
            {
                _.CALL(this, sb, "AppendLine", _.ARGS.Val(_.CONCAT("<div class=\"CalKey\">", strcalkey, "</div>")));
            }
            return BookingUI_RenderAvailCalKey_retVal;
        }

        public object bookingui_renderavailcallinks(ref object dstart, ref object sb)
        {
            object BookingUI_RenderAvailCalLinks_retVal = null;
            object dcalstartprev = null;
            object strtitleprev = null;
            object dcalstartnext = null;
            object strtitlenext = null;
            object ipositivemonthadjustment = null;
            object inegativemonthadjustment = null;

            // dStart is the start date for the last month shown in the rendered calendars
            // and we therefore only need to go forward by 1 month
            // even if no calendars are shown for the current month we can still potentially
            // move to a future month where there is availability.
            ipositivemonthadjustment = (Int16)1;
            // The previous month link has to go back by however many months are already showing, i.e. Jul & Aug are shown
            // dStart = 01/08/2011 (Aug) and we need to display Jun & Jul so we need to jump back 2 months to June.
            inegativemonthadjustment = _.SUBT(_outer.g_inumberofcalendarsrendered);

            if (_.IF(_.EQ(_.NullableNUM(_outer.g_inumberofcalendarsrendered), (Int16)0)))
            {
                // If we have no rendered calendars we still need the link to go back by 1 month
                inegativemonthadjustment = (Int16)(-1);
            }

            dcalstartprev = _.VAL(_.CALL(this, _outer.page, "Functions", "Dates", "fn_GetFirstDateOfMonth", _.ARGS.Val(_.DATEADD("m", inegativemonthadjustment, dstart))));
            strtitleprev = _.VAL(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/availcalendar/previousmonth").Val("&lt;&lt; Previous Month")));

            dcalstartnext = _.VAL(_.CALL(this, _outer.page, "Functions", "Dates", "fn_GetFirstDateOfMonth", _.ARGS.Val(_.DATEADD("m", ipositivemonthadjustment, dstart))));
            strtitlenext = _.VAL(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/availcalendar/nextmonth").Val("Next Month &gt;&gt;")));

            _.CALL(this, sb, "AppendLine", _.ARGS.Val("<div class=\"CalNavLinks\">"));
            _.CALL(this, sb, "AppendLine", _.ARGS.Val(_.CALL(this, _outer, "BookingUI_RenderAvailCalLink", _.ARGS.Ref(dcalstartprev, v55 => { dcalstartprev = v55; }).Ref(strtitleprev, v56 => { strtitleprev = v56; }).Val("prev"))));
            _.CALL(this, sb, "AppendLine", _.ARGS.Val(_.CALL(this, _outer, "BookingUI_RenderAvailCalLink", _.ARGS.Ref(dcalstartnext, v57 => { dcalstartnext = v57; }).Ref(strtitlenext, v58 => { strtitlenext = v58; }).Val("next"))));
            _.CALL(this, sb, "AppendLine", _.ARGS.Val("</div>"));

            return BookingUI_RenderAvailCalLinks_retVal;
        }

        public object bookingui_renderavailcallink(ref object dcalstartdate, ref object strtitle, ref object strclass)
        {
            object BookingUI_RenderAvailCalLink_retVal = null;
            object itm = null;
            object svalue = null;
            object strlink = null;
            object bfound = null;

            bfound = false;

            var enumerationContent3 = _.ENUMERABLE(_.CALL(this, _outer.request, "QueryString")).GetEnumerator();
            while (true)
            {
                if (!enumerationContent3.MoveNext())
                    break;
                itm = enumerationContent3.Current;
                if (_.IF(_.EQ(_.NullableSTR(itm), "isostartdate")))
                {
                    //reset date
                    object byrefalias15 = dcalstartdate;
                    try
                    {
                        svalue = _.VAL(_.CALL(this, _outer.page, "Functions", "Dates", "ISODate", _.ARGS.Ref(byrefalias15, v59 => { byrefalias15 = v59; })));
                    }
                    finally { dcalstartdate = byrefalias15; }
                    bfound = true;
                }
                else
                {
                    svalue = _.VAL(_.CALL(this, _outer.request, "QueryString", _.ARGS.Ref(itm, v60 => { itm = v60; })));
                }
                strlink = _.CONCAT(strlink, "&amp;", itm, "=", _.CALL(this, _outer.server, "UrlEncode", _.ARGS.Ref(svalue, v61 => { svalue = v61; })));
            }

            if (_.IF(_.NOT(bfound)))
            {
                object byrefalias16 = dcalstartdate;
                try
                {
                    strlink = _.CONCAT(strlink, "&amp;isostartdate=", _.CALL(this, _outer.server, "UrlEncode", _.ARGS.Val(_.CALL(this, _outer.page, "Functions", "Dates", "ISODate", _.ARGS.Ref(byrefalias16, v63 => { byrefalias16 = v63; })))));
                }
                finally { dcalstartdate = byrefalias16; }
            }

            if (_.IF(_.NOTEQ(_.NullableSTR(_.TRIM(_.CONCAT("", strlink))), "")))
            {
                strlink = _.REPLACE(strlink, "&amp;", "?", (Int16)1, (Int16)1, (Int16)0);
            }

            if (_.IF(_.GTE(_.NullableNUM(_.DATEDIFF("m", _.DATE(), dcalstartdate)), (Int16)0)))
            {
                BookingUI_RenderAvailCalLink_retVal = _.CONCAT("<a href=\"", strlink, "\" class=\"", strclass, "\" title=\"", strtitle, "\" rel=\"nofollow\">", strtitle, "</a>", VBScriptConstants.vbCrLf);
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
        public object bookingui_staymain_polling(ref object objdata, ref object objrendersettings)
        {
            object BookingUI_StayMain_Polling_retVal = null;
            object po = null;
            object dstartnight = null;
            object inights = null;
            object objavail = null;
            object intprodkey = null;
            object bistelebooking = null;
            object objavailentry = null;
            object bnoresults = null;
            object brenderedsummary = null;
            object intindex = null;
            object intindexsupplier = null;
            object objfuzzystayoptions = null;
            object objfuzzystay = null;
            object bprecisematch = null;
            object bstayhaslocalavail = null;
            object objsuppliersforstay = null;
            object objsupplier = null;
            object objdictavaistays = null;
            object stravailstaykey = null;
            object arystay = null;
            object sstayno = null;
            object bstayindicative = null;
            object brenderedinitialstay = null;
            object istaynum = null;
            object reqdicttemp = null;
            object bookingtype = null; /* Undeclared in source */

            po = _.OBJ(_.CALL(this, objrendersettings, "OutputWriter"));
            dstartnight = _.VAL(_.CALL(this, objrendersettings, "BookingRequirement", "VisitDate"));
            inights = _.VAL(_.CALL(this, objrendersettings, "BookingRequirement", "Nights"));

            // This is new, VB Polling approach (only supports accommodation, but handles results from
            // multiple providers)
            objavail = _.OBJ(_.CALL(this, objdata, "Availability"));
            intprodkey = _.VAL(_.CALL(this, objdata, "Product_Key"));
            bistelebooking = _.VAL(_.CALL(this, objdata, "IsOnTeleBookingChannel"));

            objdictavaistays = _.OBJ(_.CALL(this, _outer.server, "CreateObject", _.ARGS.Val("Scripting.Dictionary")));

            // Quick situation assertion
            if (_.IF(_.NOTEQ(_.NullableSTR(_.CALL(this, objrendersettings, "BookingType")), "accommodation")))
            {
                _.RAISEERROR(VBScriptConstants.vbObjectError, "ETWP.BookingUnitSelection", _.CONCAT("BookingUI_StayMain_Polling: BookingType not supported (\"", bookingtype, "\")"));
            }
            if (_.IF(_.GT(_.NullableNUM(_.CALL(this, objrendersettings, "BookingRequirement", "Offer")), (Int16)0)))
            {
                _.RAISEERROR(VBScriptConstants.vbObjectError, "ETWP.BookingUnitSelection", _.CONCAT("BookingUI_StayMain_Polling: Not supported with Conference Bookings (OfferKey = ", _.CALL(this, objrendersettings, "BookingRequirement", "Offer"), ")"));
            }

            // Grab hold of the data for the stay(s) - ensure we've got some availability
            objfuzzystayoptions = _.OBJ(_.CALL(this, objavail, "GetUniqueFuzzyCombinations", _.ARGS.ForceBrackets()));
            if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, objfuzzystayoptions, "Count")), (Int16)0)))
            {
                _.CALL(this, _outer.page, "PrintTraceWarning", _.ARGS.Val("objAvail.GetUniqueFuzzyCombinations reported zero stay options"));
                bnoresults = true;
            }
            else
            {
                // Double-check that all stay options report availability - there shouldn't be any stay
                // data returned that doesn't have avail data

                bnoresults = false;
                var loopEnd10 = _.NUM(_.SUBT(_.CALL(this, objfuzzystayoptions, "Count"), (Int16)1));
                var loopStart10 = _.NUM((Int16)0, loopEnd10, (Int16)1);
                if (_.StrictLTE(loopStart10, loopEnd10))
                {
                    for (intindex = loopStart10; _.StrictLTE(intindex, loopEnd10); intindex = _.ADD(intindex, (Int16)1))
                    {
                        objfuzzystay = _.OBJ(_.CALL(this, objfuzzystayoptions, "GetItem", _.ARGS.Ref(intindex, v65 => { intindex = v65; })));
                        objsuppliersforstay = _.OBJ(_.CALL(this, objavail, "GetSupplierUnitDataForStay", _.ARGS.Val(_.CALL(this, objfuzzystay, "StartDate")).Val(_.CALL(this, objfuzzystay, "Nights"))));
                        if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, objsuppliersforstay, "Count")), (Int16)0)))
                        {
                            _.CALL(this, _outer.page, "PrintTraceWarning", _.ARGS.Val(_.CONCAT("Stay (", _.CALL(this, objfuzzystay, "StartDate"), ", ", _.CALL(this, objfuzzystay, "Nights"), ") reported zero suppliers")));
                            bnoresults = true;
                        }
                        else
                        {
                            var loopEnd11 = _.NUM(_.SUBT(_.CALL(this, objsuppliersforstay, "Count"), (Int16)1));
                            var loopStart11 = _.NUM((Int16)0, loopEnd11, (Int16)1);
                            if (_.StrictLTE(loopStart11, loopEnd11))
                            {
                                for (intindexsupplier = loopStart11; _.StrictLTE(intindexsupplier, loopEnd11); intindexsupplier = _.ADD(intindexsupplier, (Int16)1))
                                {
                                    objsupplier = _.OBJ(_.CALL(this, objsuppliersforstay, "GetItem", _.ARGS.Ref(intindexsupplier, v66 => { intindexsupplier = v66; })));
                                    if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, objsupplier, "Units", "Count")), (Int16)0)))
                                    {
                                        _.CALL(this, _outer.page, "PrintTraceWarning", _.ARGS.Val(_.CONCAT("Supplier ", _.CALL(this, objsupplier, "Name"), " for Stay (", _.CALL(this, objfuzzystay, "StartDate"), ", ", _.CALL(this, objfuzzystay, "Nights"), ") reported zero units")));
                                        bnoresults = true;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // If not, render error and get out
            if (_.IF(bnoresults))
            {
                // Render message, set ProdHasAvail To False (only
                // used by BookingKeys control, I think) and close recordsets
                object byrefalias17 = objrendersettings;
                try
                {
                    _.CALL(this, _outer, "RenderNoAvailElement", _.ARGS.Ref(byrefalias17, v67 => { byrefalias17 = v67; }));
                }
                finally { objrendersettings = byrefalias17; }
                _outer.bprodhasavail = false; // This is exposed through the WSC's public property "ProdHasAvail"
                return BookingUI_StayMain_Polling_retVal;
            }

            _outer.bprodhasavail = true; // This is exposed through the WSC's public property "ProdHasAvail"

            // Loop through different stay options
            // - Store data for all stays for calendar
            if (_.IF(_outer.brenderascalendar))
            {

                var loopEnd12 = _.NUM(_.SUBT(_.CALL(this, objfuzzystayoptions, "Count"), (Int16)1));
                var loopStart12 = _.NUM((Int16)0, loopEnd12, (Int16)1);
                if (_.StrictLTE(loopStart12, loopEnd12))
                {
                    for (intindex = loopStart12; _.StrictLTE(intindex, loopEnd12); intindex = _.ADD(intindex, (Int16)1))
                    {
                        objfuzzystay = _.OBJ(_.CALL(this, objfuzzystayoptions, "GetItem", _.ARGS.Ref(intindex, v68 => { intindex = v68; })));

                        stravailstaykey = _.CONCAT("sd_", _.CALL(this, objfuzzystay, "StartDate"));
                        if (_.IF(_.CALL(this, objdictavaistays, "Exists", _.ARGS.Ref(stravailstaykey, v69 => { stravailstaykey = v69; }))))
                        {

                            // We expect value in the format [stayNo]_[indicative]
                            arystay = _.SPLIT(_.CALL(this, objdictavaistays, _.ARGS.Ref(stravailstaykey, v70 => { stravailstaykey = v70; })), "_");
                            sstayno = _.VAL(_.CALL(this, arystay, _.ARGS.Val((Int16)0)));
                            sstayno = _.CONCAT(sstayno, "-", intindex);

                            bstayindicative = _.CBOOL(_.CALL(this, arystay, _.ARGS.Val((Int16)1)));
                            if (_.IF(_.AND(_.NOT(bstayindicative), _.CALL(this, objfuzzystay, "Indicative"))))
                            {
                                bstayindicative = _.VAL(_.CALL(this, objfuzzystay, "Indicative"));
                            }
                            _.SET(_.CONCAT(sstayno, "_", bstayindicative), this, objdictavaistays, null, _.ARGS.Ref(stravailstaykey, v72 => { stravailstaykey = v72; }));
                            _.ERASE(arystay, v73 => { arystay = v73; });
                        }
                        else
                        {
                            _.CALL(this, objdictavaistays, "Add", _.ARGS.Val(_.CONCAT("sd_", _.CALL(this, objfuzzystay, "StartDate"))).Val(_.CONCAT(_.ADD(intindex, (Int16)1), "_", _.CALL(this, objfuzzystay, "Indicative"))));
                        }

                    }
                }

            }

            brenderedinitialstay = false;
            // - For unit selections: If we have a perfect match stay, don't bother with the fuzzy options
            var loopEnd13 = _.NUM(_.SUBT(_.CALL(this, objfuzzystayoptions, "Count"), (Int16)1));
            var loopStart13 = _.NUM((Int16)0, loopEnd13, (Int16)1);
            if (_.StrictLTE(loopStart13, loopEnd13))
            {
                for (intindex = loopStart13; _.StrictLTE(intindex, loopEnd13); intindex = _.ADD(intindex, (Int16)1))
                {
                    // 2010-11-03 TB: Stay numbers are 1-based so add 1 to zero-based index
                    istaynum = _.ADD(intindex, (Int16)1);

                    objfuzzystay = _.OBJ(_.CALL(this, objfuzzystayoptions, "GetItem", _.ARGS.Ref(intindex, v74 => { intindex = v74; })));

                    // 2010-01-29 DWR: Need to use DateValue here since dStartNight might be a string
                    // which will cause the comparison to fail when they represent the same date
                    bprecisematch = _.AND(_.EQ(_.DATEVALUE(_.CALL(this, objfuzzystay, "StartDate")), _.DATEVALUE(dstartnight)), _.EQ(_.CALL(this, objfuzzystay, "Nights"), inights));

                    // 2010-01-29 DWR: In cases where we have a precise match and we're not rendering a calendar
                    // then we want to just display that stay and get out! If we DON'T have a precise match and
                    // we're not using the calendar approach then we want to render all options and have client
                    // side script juggle them. If we ARE rendering the calendar then we want to display ALL
                    // stays - regardless of whether we have a precise match - because the calendar relies
                    // on the data being in the markup for it to swap around.
                    if (_.IF(_.OR(_outer.brenderascalendar, _.NOT(bprecisematch))))
                    {
                        _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<div class=\"PollingFuzzySetWrapper\" id=\"stay_", istaynum, "\">")));
                    }

                    // we only render the first stay when initially loading the unitselection
                    // we get the rest via a partial render request and the data is returned as JSON
                    // ready for manipulation and insertion by javascript
                    // this is done to avoid large amounts of HTML being rendered and then hidden
                    if (_.IF(_.NOT(brenderedinitialstay)))
                    {
                        object byrefalias18 = objrendersettings;
                        try
                        {
                            _.CALL(this, _outer, "RenderStay", _.ARGS.Ref(objfuzzystay, v75 => { objfuzzystay = v75; }).Ref(objavail, v76 => { objavail = v76; }).Ref(istaynum, v77 => { istaynum = v77; }).Ref(byrefalias18, v78 => { byrefalias18 = v78; }).Ref(bistelebooking, v79 => { bistelebooking = v79; }).Val(_.CALL(this, objdata, "bookingweb")).Val(_.CALL(this, objdata, "EviivoId")).Val(_.CALL(this, objdata, "Units")));
                        }
                        finally { objrendersettings = byrefalias18; }
                    }

                    if (_.IF(_outer.brenderascalendar))
                    {
                        brenderedinitialstay = true;
                    }

                    // Close the wrapper for the current stay date/length result set
                    // 2010-01-29: See earlier comment about this..
                    if (_.IF(_.OR(_outer.brenderascalendar, _.NOT(bprecisematch))))
                    {
                        _.CALL(this, po, "Write", _.ARGS.Val("</div>"));
                    }

                    // If these options were a perfect match, drop out
                    // 2010-01-29 DWR: Unless we're rendering the calendar! In this case client-side javascript
                    // will look after showing one fuzzy stay at a time, but it needs all data present.
                    if (_.IF(_.AND(bprecisematch, _.NOT(_outer.brenderascalendar))))
                    {
                        break;
                    }

                }
            }

            if (_.IF(_outer.brenderascalendar))
            {

                reqdicttemp = _.OBJ(_.CALL(this, _outer.page, "Functions", "GetNewObject", _.ARGS.Val("RequestDict")));
                _.CALL(this, reqdicttemp, "ForceAdd", _.ARGS.Val("AsyncAction").Val("unitselection"));
                _.CALL(this, reqdicttemp, "ForceAdd", _.ARGS.Val("PartialRenderControlList").Val(_.CALL(this, _outer.context, "PageControlKey")));
                _.CALL(this, reqdicttemp, "ForceAdd", _.ARGS.Val("Silent").Val("1"));
                _.CALL(this, reqdicttemp, "Remove", _.ARGS.Val("Debug"));
                _.CALL(this, reqdicttemp, "Remove", _.ARGS.Val("PartialRenderType"));
                _.CALL(this, reqdicttemp, "Remove", _.ARGS.Val("Trace"));

                _.CALL(this, _outer.page, "PrintTrace", _.ARGS.Val("BookingUI_StayMain_Polling: Render available stays as calendars - start"));
                _.CALL(this, _outer, "BookingUI_RenderAvailCal", _.ARGS.Ref(po, v80 => { po = v80; }).Ref(objdictavaistays, v81 => { objdictavaistays = v81; }).Val(false));
                _.CALL(this, _outer.page, "PrintTrace", _.ARGS.Val("BookingUI_StayMain_Polling: Render available stays as calendars - end"));
                _.CALL(this, po, "Write", _.ARGS.Val("<script type=\"text/javascript\">"));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("NewMind.ETWP.ControlData[", _.CALL(this, _outer.context, "PageControlKey"), "] = { ")));
                _.CALL(this, po, "Write", _.ARGS.Val("UnitSelPartialRenderLink: '"));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _.CALL(this, _outer.page, "Functions", "GetNewObject", _.ARGS.Val("clsJSON")), "EscapeJSON", _.ARGS.Val(_.CONCAT("?", _.CALL(this, reqdicttemp, "Querystring"))))));
                _.CALL(this, po, "Write", _.ARGS.Val("'"));
                _.CALL(this, po, "Write", _.ARGS.Val(" };"));
                _.CALL(this, po, "Write", _.ARGS.Val("NewMind.ETWP.Booking.InitUnitSel();"));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("</", "script>")));

            }

            // Kick off the show / hide script for fuzzy result sets now that we've rendered out all
            // the content rather than waiting for page load - hopefully we can remove some of the
            // flicker that occurs otherwise
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<script type=\"text/javascript\">NewMind.ETWP.Booking.InitPollingUnitSel();</", "script>")));

            return BookingUI_StayMain_Polling_retVal;
        }

        public object rendernotrequireddatewarning(ref object po)
        {
            object RenderNotRequiredDateWarning_retVal = null;
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<p class=\"fuzzyWarning\">", _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/notrequireddates").Val("Sorry, we don't have any availability for the dates you requested. These are the nearest available dates for your room and duration requirements.")), "</p>")));
            return RenderNotRequiredDateWarning_retVal;
        }

        public object renderstay(ref object objfuzzystay, ref object objavail, ref object intindex, ref object objrendersettings, ref object bistelebooking, object strproductbookingwebifany, object streviivoidifany, object objallunits)
        {
            object RenderStay_retVal = null;
            object objsuppliersforstay = null;
            object intprodkey = null;
            object dstartnight = null;
            object inights = null;
            object po = null;
            object bstayhaslocalavail = null;
            object intextsuppliersshown = null;
            object brenderedstaysummary = null;
            object intindexsupplier = null;
            object objsupplier = null;
            object bskipsupplier = null;
            object strbookingstaysummary = null;
            object bexternalsupplier = null;
            object strsupplierid = null;
            object strsuppliername = null;
            object strsupplierquality = null;
            object strsupplierlogo = null;
            object strsuppliereviivoname = null;
            object intbookingtype = null;
            object bprecisematch = null;

            // 2011-08-09 DWR: Expect the BookingRequirement in objRenderSettings to be read-only (since it usually comes from Page.Functions.GetSharedObject),
            // so replace it with an editable version (since some methods in here try to mess about with properties on it)
            _.SET(_.OBJ(_.CALL(this, _outer, "GetEditableBookingRequirement", _.ARGS.Val(_.CALL(this, objrendersettings, "BookingRequirement")))), this, objrendersettings, "BookingRequirement");

            objsuppliersforstay = _.OBJ(_.CALL(this, objavail, "GetSupplierUnitDataForStay", _.ARGS.Val(_.CALL(this, objfuzzystay, "StartDate")).Val(_.CALL(this, objfuzzystay, "Nights"))));

            intprodkey = _.VAL(_.CALL(this, objrendersettings, "ProductKey"));
            dstartnight = _.VAL(_.CALL(this, objrendersettings, "BookingRequirement", "VisitDate"));
            inights = _.VAL(_.CALL(this, objrendersettings, "BookingRequirement", "Nights"));
            po = _.OBJ(_.CALL(this, objrendersettings, "OutputWriter"));

            //just need to set these here as we may be coming in direct from partial render request
            _outer.brenderascalendar = _.VAL(_.CALL(this, objrendersettings, "RenderAsCalendar"));
            _outer.isvbpollingenabled = _.VAL(_.CALL(this, objrendersettings, "IsVBPollingEnabled"));

            // Loop through each supplier and render their units
            // - Suppliers will be ordered NewMind, FrontDesk, Other
            // - If "Booking_ForceExternal" is enabled, FrontDesk is treated as "Other"
            // - There is a limit on the number of "Other" entries to be rendered (if ForceExternal
            //   is enabled, then FrontDesk counts towards this limit)
            // - If ForceExternal is not enabled, FrontDesk will only be rendered if there is no
            //   local availability
            bstayhaslocalavail = false;
            intextsuppliersshown = (Int16)0;
            brenderedstaysummary = false;

            bprecisematch = _.AND(_.EQ(_.DATEVALUE(_.CALL(this, objfuzzystay, "StartDate")), _.DATEVALUE(dstartnight)), _.EQ(_.CALL(this, objfuzzystay, "Nights"), inights));

            var loopEnd14 = _.NUM(_.SUBT(_.CALL(this, objsuppliersforstay, "Count"), (Int16)1));
            var loopStart14 = _.NUM((Int16)0, loopEnd14, (Int16)1);
            if (_.StrictLTE(loopStart14, loopEnd14))
            {
                for (intindexsupplier = loopStart14; _.StrictLTE(intindexsupplier, loopEnd14); intindexsupplier = _.ADD(intindexsupplier, (Int16)1))
                {

                    // Get basic supplier data - count FrontDesk as "Other" if ForceExternal enabled
                    objsupplier = _.OBJ(_.CALL(this, objsuppliersforstay, "GetItem", _.ARGS.Ref(intindexsupplier, v82 => { intindexsupplier = v82; })));
                    if (_.IF(_.CALL(this, objsupplier, "IsLocal")))
                    {
                        bstayhaslocalavail = true;
                    }
                    bexternalsupplier = _.VAL(_.OR(_.CALL(this, objsupplier, "IsExternal"), _.AND(_.CALL(this, objsupplier, "IsRemote"), _.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("Booking_ForceExternal")))));

                    // Don't render FrontDesk if got local avail for this stay and not enabled ForceExternal
                    bskipsupplier = _.VAL(_.AND(bstayhaslocalavail, _.AND(_.CALL(this, objsupplier, "IsRemote"), _.NOT(_.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("Booking_ForceExternal"))))));
                    if (_.IF(_.NOT(bskipsupplier)))
                    {

                        // Don't bother rendering stay summary title if we've got a perfect match, as we
                        // won't be showing any fuzzy content if there's a spot-on option
                        if (_.IF(_.AND(_.NOT(bprecisematch), _.NOT(brenderedstaysummary))))
                        {
                            if (_.IF(_.NOT(_outer.brenderascalendar)))
                            {
                                _.CALL(this, _outer, "BookingUI_StaySummary", _.ARGS.Ref(dstartnight, v83 => { dstartnight = v83; }).Ref(inights, v84 => { inights = v84; }).Val(_.CALL(this, objfuzzystay, "StartDate")).Val(_.CALL(this, objfuzzystay, "Nights")).Ref(po, v85 => { po = v85; }));
                            }
                            brenderedstaysummary = true;
                        }

                        // If this is an external supplier, we need the deep-link quality to pass to get
                        // included in the hidden booking-info form fields

                        // PW 2010-07-28 I have added a new field called strSupplierEviivoName
                        // This is to pass through the original name field from Eviivo through to the polling exit page.
                        // Previously, we did some manipulation on this value to ensure it had a nice display name.
                        // However, this had broken Eviivo's own external link - we have in the past asked Eviivo to provide a
                        // nice display name field but until they do so we are going to have to do our own and pass both values through as hidden
                        // form fields
                        if (_.IF(bexternalsupplier))
                        {
                            strsupplierid = _.VAL(_.CALL(this, objsupplier, "ID"));
                            strsuppliername = _.VAL(_.CALL(this, objsupplier, "DisplayName"));
                            strsupplierquality = _.VAL(_.CALL(this, objsupplier, "Quality"));
                            strsuppliereviivoname = _.VAL(_.CALL(this, objsupplier, "Name"));
                        }
                        else
                        {
                            strsupplierid = VBScriptConstants.Null;
                            strsuppliername = VBScriptConstants.Null;
                            strsupplierquality = VBScriptConstants.Null;
                            strsuppliereviivoname = VBScriptConstants.Null;
                        }

                        // Render the actual options (wrap in the standard form tag)
                        if (_.IF(_.CALL(this, objsupplier, "IsLocal")))
                        {
                            if (_.IF(_.ISEMPTY(_outer.isexternalbooking)))
                            {
                                _.CALL(this, _outer, "InitExternalBookingSettings", _.ARGS.ForceBrackets());
                            }
                            if (_.IF(_outer.isexternalbooking))
                            {
                                intbookingtype = _.VAL(_outer.booking_redirect);
                                _outer.strproductestateid = _.VAL(_.CALL(this, _outer.dms, "GetProductEstateID", _.ARGS.Ref(intprodkey, v86 => { intprodkey = v86; })));
                                _outer.strextbookurl = _.VAL(_.CALL(this, _outer, "GetExtBookUrlFromProductEstate", _.ARGS.Ref(_outer.strproductestateid, v87 => { _outer.strproductestateid = v87; })));
                            }
                            else
                            {
                                intbookingtype = _.VAL(_outer.booking_local);
                            }
                        }
                        else if (_.IF(_.CALL(this, objsupplier, "IsExternal")))
                        {
                            // 2011-07-20 DWR: We don't need to call InitExternalBookingSettings if dealing with an VB Polling product as
                            // the next page should always be the Polling Exit (no point redirecting to another site which will then - if
                            // it's an NM site - have to display another redirect page to book the product)
                            intbookingtype = _.VAL(_outer.booking_pollingredirect);
                        }
                        else
                        {
                            if (_.IF(_.ISEMPTY(_outer.isexternalbooking)))
                            {
                                _.CALL(this, _outer, "InitExternalBookingSettings", _.ARGS.ForceBrackets());
                            }
                            if (_.IF(_outer.isexternalbooking))
                            {
                                intbookingtype = _.VAL(_outer.booking_redirect);
                                _outer.strproductestateid = _.VAL(_.CALL(this, _outer.dms, "GetProductEstateID", _.ARGS.Ref(intprodkey, v88 => { intprodkey = v88; })));
                                _outer.strextbookurl = _.VAL(_.CALL(this, _outer, "GetExtBookUrlFromProductEstate", _.ARGS.Ref(_outer.strproductestateid, v89 => { _outer.strproductestateid = v89; })));
                            }
                            else
                            {
                                intbookingtype = _.VAL(_outer.booking_eviivo);
                            }
                        }

                        // Local and FrontDesk both use current site name w/out logo
                        // External Suppliers should have their own logo passed in
                        // PW - 	moved this out of BookingUI_StayDetails_PollingHeader
                        //		we can now pass it to the hidden form fields
                        //		for use on the polling exit page
                        if (_.IF(_.CALL(this, objsupplier, "IsExternal")))
                        {
                            strsupplierlogo = _.VAL(_.CALL(this, objsupplier, "Logo"));
                            if (_.IF(_.EQ(_.NullableSTR(_.TRIM(_.CONCAT("", strsuppliername))), "")))
                            {
                                strsuppliername = "Unnamed Supplier";
                            }
                            else if (_.IF(_.EQ(_.NullableSTR(strsupplierlogo), "")))
                            {
                                // 2014-07-01 DWR: It's common for Eviivo to not return logo data for the Polling Providers so for most cases we take the Supplier Name (the
                                // Eviivo version, rather than the "friendly" version that we maintain) and request the logo from ntop using it. For cases where Eviivo
                                // results are treated as Polling results (see FogBugz 10386), we need a special case (the friendly name will always be "Eviivo" in
                                // this case).
                                if (_.IF(_.EQ(_.NullableSTR(strsuppliername), "Eviivo")))
                                {
                                    strsupplierlogo = _.VAL(_.CALL(this, _outer.page, "ImageResource", _.ARGS.Val("bookonline/unitselection/polling/eviivo").Val("/engine/shared_gfx/eviiopollingresult.jpg")));
                                }
                                else
                                {
                                    // 2008-12-09 DWR: Supplier Logo isn't actually going to be received from the Eviivo Component, we mash Supplier Name into this url
                                    // 2010-03-04 DWR: Eviivo moved the logo location..
                                    strsupplierlogo = _.CONCAT("http://www.ntopsearch.com/media/images/Suppliers/", strsuppliereviivoname, ".gif");
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
                            strsuppliername = "";
                            if (_.IF(_outer.isexternalbooking))
                            {
                                strsuppliername = _.VAL(_.CALL(this, _outer.page, "Resource", _.ARGS.Val(_.CONCAT("bookonline/unitselection/polling/localsupplier/estate_", _outer.strproductestateid, "/name")).Val("")));
                            }
                            if (_.IF(_.EQ(_.NullableSTR(strsuppliername), "")))
                            {
                                //#MJ -	the resource manage is the same for both main sites and channel sites
                                //		therefore we can never use Page.Site.Name as an alternative value as this would be cached wrongly by the ResourceManager
                                //		so try to pull one from there, if not fall back to the site name
                                strsuppliername = _.VAL(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/polling/localsupplier/name").Val("")));
                                if (_.IF(_.EQ(_.NullableSTR(strsuppliername), "")))
                                {
                                    strsuppliername = _.VAL(_.CALL(this, _outer.page, "Site", "Name"));
                                }
                            }
                            // - Supplier Logo
                            strsupplierlogo = _.VAL(_.CALL(this, _outer, "GetSupplierLogo", _.ARGS.Ref(_outer.strproductestateid, v90 => { _outer.strproductestateid = v90; })));

                        }

                        // 2013-02-05 TB: objRenderSettings is used by RenderBookingInfoForm to populate some hidden stay information
                        // For fuzzy stays, both nights and startdate may differ from the original requirements.
                        // For FogBugz case 7594 I added the second line below which wasn't present.
                        _.SET(_.VAL(_.CALL(this, objfuzzystay, "StartDate")), this, _.CALL(this, objrendersettings, "BookingRequirement"), "VisitDate");
                        _.SET(_.VAL(_.CALL(this, objfuzzystay, "Nights")), this, _.CALL(this, objrendersettings, "BookingRequirement"), "Nights");

                        // 2014-03-12 DWR: We need to pass the Search Industry Classification into the form rendering code for VB Polling Products so that the
                        // Polling Exist can generate the deep link correctly. An Eviivo Configset can be set up with zero, meaning support either 1 OR 9. The
                        // Avail Component will perform searches for both in that case but only allow any Products to return results for one. Since we won't
                        // get an objSupplier reference with zero units (since that would mean it's not got availability and we're only looking at available
                        // options here) we can just grab the IndustryClassification values from the first Unit since it is guaranteed to be consistent
                        // across all Units for this booking option. The IndustryClassification value will be zero for non-Eviivo data but that won't
                        // matter since it's only ever consider in the Polling Exit which is for Eviivo results only.
                        object byrefalias19 = objrendersettings;
                        try
                        {
                            _.CALL(this, _outer, "RenderBookingInfoForm", _.ARGS.Ref(po, v91 => { po = v91; }).Ref(intprodkey, v92 => { intprodkey = v92; }).Ref(byrefalias19, v93 => { byrefalias19 = v93; }).Ref(intbookingtype, v94 => { intbookingtype = v94; }).Ref(strsupplierid, v95 => { strsupplierid = v95; }).Ref(strsuppliername, v96 => { strsuppliername = v96; }).Ref(strsuppliereviivoname, v97 => { strsuppliereviivoname = v97; }).Ref(strsupplierquality, v98 => { strsupplierquality = v98; }).Ref(strsupplierlogo, v99 => { strsupplierlogo = v99; }).Val(_.CALL(this, _.CALL(this, objsupplier, "Units", "GetItem", _.ARGS.Val((Int16)0)), "IndustryClassification")));
                        }
                        finally { objrendersettings = byrefalias19; }

                        _.CALL(this, _outer, "BookingUI_StayDetails_PollingHeader", _.ARGS.Ref(objsupplier, v100 => { objsupplier = v100; }).Ref(po, v101 => { po = v101; }).Ref(strsupplierlogo, v102 => { strsupplierlogo = v102; }).Ref(strsuppliername, v103 => { strsuppliername = v103; }));

                        // 2009-09-14 DWR: Forcing iStayNum to "1" every time - since we are clearly only having
                        // one stay per form (since we open the form above - in RenderBookingInfoForm - and we
                        // close it below) we'll always be passing only a single stay to the next stage. This
                        // makes things easier - the multiple-stays-per-form idea was ridiculous.
                        // 2010-10-21 TB: Changing back to use unique stay index. Multiple stays per form will
                        // happen for fuzzy results and calendar view. html ids use the stay key, as does the JS
                        // when choosing to show/hide the book now button.
                        object byrefalias20 = intindex, byrefalias21 = bistelebooking;
                        try
                        {
                            _.CALL(this, _outer, "BookingUI_StayDetails", _.ARGS.Ref(objsupplier, v104 => { objsupplier = v104; }).Ref(byrefalias20, v105 => { byrefalias20 = v105; }).Ref(dstartnight, v106 => { dstartnight = v106; }).Ref(inights, v107 => { inights = v107; }).Ref(byrefalias21, v108 => { byrefalias21 = v108; }).Ref(strproductbookingwebifany, v109 => { strproductbookingwebifany = v109; }).Ref(streviivoidifany, v110 => { streviivoidifany = v110; }).Val(_.CALL(this, objrendersettings, "ProductKey")).Val(_.CALL(this, objrendersettings, "Channel")).Val(_.CALL(this, objfuzzystay, "Indicative")).Val(_.NOT(_.CALL(this, objfuzzystay, "HasInvalidIndicative"))).Ref(objallunits, v111 => { objallunits = v111; }).Val(VBScriptConstants.Nothing).Ref(po, v112 => { po = v112; }).Val(false));
                        }
                        finally { intindex = byrefalias20; bistelebooking = byrefalias21; } //we can't render the maximum available units for polling

                        _.CALL(this, po, "Write", _.ARGS.Val("</form>"));

                    }

                }
            }

            return RenderStay_retVal;
        }

        public object rendernoavailelement(object objrendersettings)
        {
            object RenderNoAvailElement_retVal = null;
            object po = null;
            object strclassmonth = null;

            po = _.OBJ(_.CALL(this, objrendersettings, "OutputWriter"));
            _.CALL(this, po, "Write", _.ARGS.Val("<div class=\"pnNoAvail\">"));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/noavailability").Val("<p>No availability for this product for the specified date. This may occur if the accommodation is booked prior to your arrival at this page.</p>"))));
            _.CALL(this, po, "Write", _.ARGS.Val("</div>"));

            if (_.IF(_.CALL(this, objrendersettings, "RenderAsCalendar")))
            {

                strclassmonth = "MonthWrapper";

                _.CALL(this, po, "Write", _.ARGS.Val("<div class=\"CalendarsWrapper\">"));
                //
                _.CALL(this, _outer, "BookingUI_RenderCalendarMonth", _.ARGS.Ref(po, v113 => { po = v113; }).Val(_.CALL(this, objrendersettings, "BookingRequirement", "VisitDate")).Val(_.CONCAT(strclassmonth, " currentmonth")));
                //					' last day + 1 to get the first day of the next month for the calendar
                _.CALL(this, _outer, "BookingUI_RenderCalendarMonth", _.ARGS.Ref(po, v114 => { po = v114; }).Val(_.ADD(_.CALL(this, _outer.page, "Functions", "Dates", "fn_GetLastDateOfMonth", _.ARGS.Val(_.CALL(this, objrendersettings, "BookingRequirement", "VisitDate"))), (Int16)1)).Val(_.CONCAT(strclassmonth, " nextmonth")));

                // global count used to track how many calendars have been added to the output for the prev/next buttons
                _outer.g_inumberofcalendarsrendered = (Int16)2;

                _.CALL(this, _outer, "BookingUI_RenderAvailCalLinks", _.ARGS.Val(_.CALL(this, objrendersettings, "BookingRequirement", "VisitDate")).Ref(po, v115 => { po = v115; }));
                _.CALL(this, _outer, "BookingUI_RenderAvailCalKey", _.ARGS.Ref(po, v116 => { po = v116; }));
                _.CALL(this, po, "Write", _.ARGS.Val("</div>"));

                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<script type=\"text/javascript\">NewMind.ETWP.Booking.UpdateCalLinks();</", "script>")));
            }

            return RenderNoAvailElement_retVal;
        }

        // ====================================================================================================
        // RENDER: Main entry point when VB Polling is disabled (or handling tickets, not acco products)
        // ====================================================================================================
        public object bookingui_staymain_legacy(ref object objdata, ref object objrendersettings)
        {
            object BookingUI_StayMain_Legacy_retVal = null;
            object po = null;
            object intbookingtype = null;
            object bnoresults = null;
            object objfuzzystayoptions = null;
            object objfuzzystay = null;
            object objsuppliersforstay = null;
            object objavailentry = null;
            object lsremoteunitselections = null;
            object objavail = null;
            object intprodkey = null;
            object bistelebooking = null;
            // This is the non-VB-Polling approach (supports EITHER FrontDesk OR local availability for accommodation)
            //reset the output variable to our OutputWriter
            po = _.OBJ(_.CALL(this, objrendersettings, "OutputWriter"));

            objavail = _.OBJ(_.CALL(this, objdata, "Availability"));
            intprodkey = _.VAL(_.CALL(this, objdata, "Product_Key"));

            bistelebooking = _.VAL(_.CALL(this, objdata, "IsOnTeleBookingChannel"));

            // Grab hold of the data (in this method, there should only ever be zero or one fuzzy
            // stay options, as the BookingUI_StayMain_Legacy method handle fuzzy availability)
            objfuzzystayoptions = _.OBJ(_.CALL(this, objavail, "GetUniqueFuzzyCombinations", _.ARGS.ForceBrackets()));
            if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, objfuzzystayoptions, "Count")), (Int16)0)))
            {
                _.CALL(this, _outer.page, "PrintTraceWarning", _.ARGS.Val("objAvail.GetUniqueFuzzyCombinations reported zero stay options"));
                bnoresults = true;
            }
            else
            {
                // Any suppliers returned here will be sorted with Local / NewMind first, then FrontDesk
                // second (if we have both) - if there are multiple, it should always be the first one
                // that we want
                objfuzzystay = _.OBJ(_.CALL(this, objfuzzystayoptions, "GetItem", _.ARGS.Val((Int16)0)));
                _.CALL(this, _outer.page, "PrintTrace", _.ARGS.Val(_.CONCAT("BookingUI_StayMain_Legacy: Get data for stay - ", _.CALL(this, objfuzzystay, "StartDate"), ", ", _.CALL(this, objfuzzystay, "Nights"))));
                objsuppliersforstay = _.OBJ(_.CALL(this, objavail, "GetSupplierUnitDataForStay", _.ARGS.Val(_.CALL(this, objfuzzystay, "StartDate")).Val(_.CALL(this, objfuzzystay, "Nights"))));
                if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, objsuppliersforstay, "Count")), (Int16)0)))
                {
                    _.CALL(this, _outer.page, "PrintTraceWarning", _.ARGS.Val("objAvail.GetSupplierUnitDataForStay reported zero suppliers"));
                    bnoresults = true;
                }
                else
                {
                    objavailentry = _.OBJ(_.CALL(this, objsuppliersforstay, "GetItem", _.ARGS.Val((Int16)0)));
                    bnoresults = false;
                }
            }

            // Open form and prepare to wrap content in "staySelection" container
            if (_.IF(_outer.isexternalbooking))
            {
                intbookingtype = _.VAL(_outer.booking_redirect);
            }
            else
            {
                intbookingtype = _.VAL(_outer.booking_local);
            }

            object byrefalias22 = objrendersettings;
            try
            {
                _.CALL(this, _outer, "RenderBookingInfoForm", _.ARGS.Ref(po, v117 => { po = v117; }).Ref(intprodkey, v118 => { intprodkey = v118; }).Ref(byrefalias22, v119 => { byrefalias22 = v119; }).Ref(intbookingtype, v120 => { intbookingtype = v120; }).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null).Val(VBScriptConstants.Null));
            }
            finally { objrendersettings = byrefalias22; }

            _.CALL(this, po, "Write", _.ARGS.Val("<div class=\"staySelection\">"));

            // Render info (or display warning if no availability)
            if (_.IF(bnoresults))
            {
                object byrefalias23 = objrendersettings;
                try
                {
                    _.CALL(this, _outer, "RenderNoAvailElement", _.ARGS.Ref(byrefalias23, v121 => { byrefalias23 = v121; }));
                }
                finally { objrendersettings = byrefalias23; }
                _outer.bprodhasavail = false; // This is exposed through the WSC's public property "ProdHasAvail"
            }
            else
            {
                _outer.bprodhasavail = true; // This is exposed through the WSC's public property "ProdHasAvail"
                if (_.IF(_.EQ(_.NullableSTR(_.CALL(this, objrendersettings, "BookingType")), "accommodation")))
                {

                    // Retrieve any unit selections that have been passed in through the querystring
                    // - eg. when VisitBritain hooks in to complete a booking
                    // There will be an entry in lsUnitSelections for each requirement.
                    // Note that ReqNo in the avail data is one-based while the lsUnitSelections indices
                    // are zero-based, so the UnitKey for ReqNo 1 = lsUnitSelections(0). If there was no
                    // selection made for a ReqNo, the lsUnitSelections value will be zero.
                    // NB: This value might be Nothing if no selections are passed in on querystring.
                    lsremoteunitselections = _.OBJ(_.CALL(this, _outer, "BookingUI_UnitSel_GetOptionsRemoteSelected", _.ARGS.Ref(objavailentry, v122 => { objavailentry = v122; })));

                    // Render the unit selection options (pass "1" as iStayNum parameter - we'll only
                    // be rendering a single stay option here, since fuzzy isn't supported in this
                    // configuration..)
                    _.CALL(this, _outer, "BookingUI_StayDetails", _.ARGS.Ref(objavailentry, v123 => { objavailentry = v123; }).Val((Int16)1).Val(_.CALL(this, objrendersettings, "BookingRequirement", "VisitDate")).Val(_.CALL(this, objrendersettings, "BookingRequirement", "Nights")).Ref(bistelebooking, v124 => { bistelebooking = v124; }).Val(_.CALL(this, objdata, "bookingweb")).Val(_.CALL(this, objdata, "EviivoId")).Ref(intprodkey, v125 => { intprodkey = v125; }).Val(_.CALL(this, objrendersettings, "Channel")).Val(_.CALL(this, objfuzzystay, "Indicative")).Val(_.NOT(_.CALL(this, objfuzzystay, "HasInvalidIndicative"))).Val(_.CALL(this, objdata, "Units")).Ref(lsremoteunitselections, v126 => { lsremoteunitselections = v126; }).Ref(po, v127 => { po = v127; }).Val(_.CALL(this, objrendersettings, "RenderMaximumUnitsAvailable")));

                }
                else
                {
                    _.CALL(this, _outer, "BookingUI_TicketsSummary", _.ARGS.Ref(objavailentry, v128 => { objavailentry = v128; }).Val(_.CALL(this, objrendersettings, "BookingRequirement", "VisitDate")).Ref(po, v129 => { po = v129; }));
                }
            }

            // Close "staySelection" div and form
            _.CALL(this, po, "Write", _.ARGS.Val("</div>"));
            _.CALL(this, po, "Write", _.ARGS.Val("</form>"));
            return BookingUI_StayMain_Legacy_retVal;
        }

        // SUMMARY: prepare a list of UnitKey selections for each ReqNo is availability recordset
        // [rsAvail]: ADO unit recordset from availability object
        // <retval>: clsList with as many values as there are ReqNo entries, containing the UnitKey for each one
        public object bookingui_unitsel_getoptionsremoteselected(object objavailentry)
        {
            object BookingUI_UnitSel_GetOptionsRemoteSelected_retVal = null;
            object intindex = null;
            object objunit = null;
            object arrrequnitoptions = null;
            object arrrequnitselections = null;
            object intunitsel = null;
            object lsunitkeys = null;
            object bookingui_unitsel_getoptionselected = null;

            // Build up a list of unit options:
            // - Will get a list of objects where each object has properties:
            //    > ReqNo (integer)
            //    > NumPeople (integer)
            //    > Units (list of integers)
            // - We're going to loop through the availability recordset, so must remember
            //   to return it back to the beginning when we're done
            arrrequnitoptions = _.OBJ(_.CALL(this, _outer.page, "Functions", "GetNewObject", _.ARGS.Val("clsList")));
            var loopEnd15 = _.NUM(_.SUBT(_.CALL(this, objavailentry, "Units", "Count"), (Int16)1));
            var loopStart15 = _.NUM((Int16)0, loopEnd15, (Int16)1);
            if (_.StrictLTE(loopStart15, loopEnd15))
            {
                for (intindex = loopStart15; _.StrictLTE(intindex, loopEnd15); intindex = _.ADD(intindex, (Int16)1))
                {
                    objunit = _.OBJ(_.CALL(this, objavailentry, "Units", "GetItem", _.ARGS.Ref(intindex, v130 => { intindex = v130; })));
                    _.CALL(this, _outer, "BookingUI_UnitSel_AddReqUnitOption", _.ARGS.Ref(arrrequnitoptions, v131 => { arrrequnitoptions = v131; }).Val(_.CALL(this, objunit, "ReqNo")).Val(_.CALL(this, objunit, "ReqSize")).Val(_.CALL(this, objunit, "UnitKey")));
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
            intunitsel = (Int16)0;
            arrrequnitselections = _.OBJ(_.CALL(this, _outer.page, "Functions", "GetNewObject", _.ARGS.Val("clsList")));
            while (true)
            {
                intunitsel = _.ADD(intunitsel, (Int16)1);
                if (_.IF(_.GT(_.NullableNUM(_.LEN(_.CALL(this, _outer.request, _.ARGS.Val(_.CONCAT("URslt", intunitsel))))), (Int16)0)))
                {
                    _.CALL(this, _outer, "BookingUI_UnitSel_AddReqUnitSelection", _.ARGS.Ref(arrrequnitselections, v132 => { arrrequnitselections = v132; }).RefIfArray(_outer.request, _.ARGS.Val(_.CONCAT("URslt", intunitsel))).Ref(arrrequnitoptions, v133 => { arrrequnitoptions = v133; }));
                }
                else
                {
                    break;
                }
            }

            // If there were no selections passed in like this, return Nothing
            if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, arrrequnitselections, "Count")), (Int16)0)))
            {
                bookingui_unitsel_getoptionselected = VBScriptConstants.Nothing;
            }

            // Now try to return matched unit options / selections
            // - Get back a list of unit keys, one key per requirement
            //   (If failed to get a perfect match, some of these values may be zero)
            BookingUI_UnitSel_GetOptionsRemoteSelected_retVal = _.OBJ(_.CALL(this, _outer, "BookingUI_UnitSel_GetMatchedReqUnitSelection", _.ARGS.Ref(arrrequnitoptions, v134 => { arrrequnitoptions = v134; }).Ref(arrrequnitselections, v135 => { arrrequnitselections = v135; })));

            return BookingUI_UnitSel_GetOptionsRemoteSelected_retVal;
        }

        public object bookingui_unitsel_addrequnitoption(ref object arrrequnitoptions, ref object intreqno, ref object intnumpeople, ref object intunitkey)
        {
            object BookingUI_UnitSel_AddReqUnitOption_retVal = null;
            object objentry = null;
            object objentryprev = null;

            // Input list SHOULD be initialised as an empty list, but just in case..
            if (_.IF(_.OR(_.ISEMPTY(arrrequnitoptions), _.ISNULL(arrrequnitoptions))))
            {
                arrrequnitoptions = _.OBJ(_.CALL(this, _outer.page, "Functions", "GetNewObject", _.ARGS.Val("clsList")));
            }

            // If we've already got list items, check whether we're still working on the same
            // ReqNo as the previous entry. If so, add to that entry's unit list.
            if (_.IF(_.GT(_.NullableNUM(_.CALL(this, arrrequnitoptions, "Count")), (Int16)0)))
            {
                objentryprev = _.OBJ(_.CALL(this, arrrequnitoptions, _.ARGS.Val(_.SUBT(_.CALL(this, arrrequnitoptions, "Count"), (Int16)1))));
                if (_.IF(_.EQ(_.CALL(this, objentryprev, _.ARGS.Val("ReqNo")), intreqno)))
                {
                    object byrefalias24 = intunitkey;
                    try
                    {
                        _.CALL(this, _.CALL(this, objentryprev, _.ARGS.Val("Units")), "Add", _.ARGS.Ref(byrefalias24, v136 => { byrefalias24 = v136; }));
                    }
                    finally { intunitkey = byrefalias24; }
                    return BookingUI_UnitSel_AddReqUnitOption_retVal;
                }
            }

            // Need to create a new entry
            objentry = _.OBJ(_.CALL(this, _outer.page, "Functions", "GetNewObject", _.ARGS.Val("clsValueBag")));
            _.SET(_.VAL(intreqno), this, objentry, null, _.ARGS.Val("ReqNo"));
            _.SET(_.VAL(intnumpeople), this, objentry, null, _.ARGS.Val("NumPeople"));
            _.SET(_.OBJ(_.CALL(this, _outer.page, "Functions", "GetNewObject", _.ARGS.Val("clsList"))), this, objentry, null, _.ARGS.Val("Units"));
            object byrefalias25 = intunitkey;
            try
            {
                _.CALL(this, _.CALL(this, objentry, _.ARGS.Val("Units")), "Add", _.ARGS.Ref(byrefalias25, v137 => { byrefalias25 = v137; }));
            }
            finally { intunitkey = byrefalias25; }
            _.CALL(this, arrrequnitoptions, "Add", _.ARGS.Ref(objentry, v138 => { objentry = v138; }));

            return BookingUI_UnitSel_AddReqUnitOption_retVal;
        }

        public object bookingui_unitsel_addrequnitselection(ref object arrrequnitselections, ref object strunitselinfo, ref object arrrequnitoptions)
        {
            object BookingUI_UnitSel_AddReqUnitSelection_retVal = null;
            var errOn = _.GETERRORTRAPPINGTOKEN();
            object arrsegments = null;
            object intnumadults = null;
            object intnumchildren = null;
            object intunitkey = null;
            object intindex = null;
            object objentry = null;
            object objunitlist = null;

            // Input list SHOULD be initialised as an empty list, but just in case..
            bool ifResult;
            object byrefalias26 = arrrequnitselections;
            try
            {
                ifResult = _.IF(() => _.OR(_.ISEMPTY(byrefalias26), _.ISNULL(byrefalias26)), errOn);
            }
            finally { arrrequnitselections = byrefalias26; }
            if (ifResult)
            {
                object byrefalias27 = arrrequnitselections;
                try
                {
                    byrefalias27 = _.OBJ(_.CALL(this, _outer.page, "Functions", "GetNewObject", _.ARGS.Val("clsList")));
                }
                finally { arrrequnitselections = byrefalias27; }
            }

            // strUnitSelInfo should be of the form "UnitKey,NumAdults,NumChildren"
            // Exit if not
            object byrefalias28 = strunitselinfo;
            try
            {
                arrsegments = _.SPLIT(byrefalias28, ",");
            }
            finally { strunitselinfo = byrefalias28; }
            if (_.IF(_.NOTEQ(_.NullableNUM(_.UBOUND(arrsegments)), (Int16)2)))
            {
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return BookingUI_UnitSel_AddReqUnitSelection_retVal;
            }

            // Ensure entries in string are numeric (exit if not)
            _.STARTERRORTRAPPINGANDCLEARANYERROR(errOn);
            _.HANDLEERROR(errOn, () => {
                intunitkey = _.CLNG(_.CALL(this, arrsegments, _.ARGS.Val((Int16)0)));
            });
            if (_.IF(() => _.ERR, errOn))
            {
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return BookingUI_UnitSel_AddReqUnitSelection_retVal;
            }
            _.HANDLEERROR(errOn, () => {
                intnumadults = _.CLNG(_.CALL(this, arrsegments, _.ARGS.Val((Int16)1)));
            });
            if (_.IF(() => _.ERR, errOn))
            {
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return BookingUI_UnitSel_AddReqUnitSelection_retVal;
            }
            _.HANDLEERROR(errOn, () => {
                intnumchildren = _.CLNG(_.CALL(this, arrsegments, _.ARGS.Val((Int16)2)));
            });
            if (_.IF(() => _.ERR, errOn))
            {
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return BookingUI_UnitSel_AddReqUnitSelection_retVal;
            }
            _.STOPERRORTRAPPINGANDCLEARANYERROR(errOn);

            // Ensure values look reasonable
            if (_.IF(_.OR(_.OR(_.OR(_.LTE(_.NullableNUM(intunitkey), (Int16)0), _.LT(_.NullableNUM(intnumadults), (Int16)0)), _.LT(_.NullableNUM(intnumchildren), (Int16)0)), _.LTE(_.NullableNUM(_.ADD(intnumadults, intnumchildren)), (Int16)0))))
            {
                _.RELEASEERRORTRAPPINGTOKEN(errOn);
                return BookingUI_UnitSel_AddReqUnitSelection_retVal;
            }

            // Preparer new entry
            objentry = _.OBJ(_.CALL(this, _outer.page, "Functions", "GetNewObject", _.ARGS.Val("clsValueBag")));
            _.SET(_.ADD(intnumadults, intnumchildren), this, objentry, null, _.ARGS.Val("NumPeople"));
            _.SET(_.VAL(intunitkey), this, objentry, null, _.ARGS.Val("UnitKey"));
            _.SET(_.OBJ(_.CALL(this, _outer.page, "Functions", "GetNewObject", _.ARGS.Val("clsList"))), this, objentry, null, _.ARGS.Val("PossReqNos"));

            // Look through the unit options and look for possible requirement matches
            // - We've got a set of requirement / room options from the DMS and we've (possibly) got a
            //   set of unit selections from VisitBritain (or whoever), but these may not currently be
            //   aligned, so we want to determine the possible ways they MIGHT go together, and we'll
            //   try to get the best configuration (which will hopefully match the original choice)
            //   later on.
            bool ifResult2;
            object byrefalias29 = arrrequnitoptions;
            ifResult2 = _.IF(() => _.GT(_.NullableNUM(_.CALL(this, byrefalias29, "Count")), (Int16)0), errOn);
            if (ifResult2)
            {
                object loopEnd16 = 0, loopStart16 = 0;
                var loopConstraintsInitialized = false;
                object byrefalias30 = arrrequnitoptions;
                _.HANDLEERROR(errOn, () => {
                    loopEnd16 = _.NUM(_.SUBT(_.CALL(this, byrefalias30, "Count"), (Int16)1));
                    loopStart16 = _.NUM((Int16)0);
                    if ((loopStart16 is DateTime) || (loopStart16 is Decimal))
                        intindex = loopStart16;
                    loopStart16 = _.NUM((Int16)0, loopEnd16, (Int16)1);
                    loopConstraintsInitialized = true;
                });
                if (_.StrictLTE(loopStart16, loopEnd16))
                {
                    if (loopConstraintsInitialized)
                        intindex = loopStart16;
                    while (true)
                    {
                        // If requirement option matches the selection's NumPeople and contains the
                        // UnitKey, then we've got a possible match
                        bool ifResult3;
                        object byrefalias31 = arrrequnitoptions;
                        ifResult3 = _.IF(() => _.AND(_.EQ(_.CALL(this, _.CALL(this, byrefalias31, _.ARGS.Ref(intindex, v141 => { intindex = v141; })), _.ARGS.Val("NumPeople")), _.CALL(this, objentry, _.ARGS.Val("NumPeople"))), _.CALL(this, _.CALL(this, _.CALL(this, byrefalias31, _.ARGS.Ref(intindex, v142 => { intindex = v142; })), _.ARGS.Val("Units")), "Contains", _.ARGS.RefIfArray(objentry, _.ARGS.Val("UnitKey")))), errOn);
                        if (ifResult3)
                        {
                            object byrefalias32 = arrrequnitoptions;
                            _.CALL(this, _.CALL(this, objentry, _.ARGS.Val("PossReqNos")), "Add", _.ARGS.RefIfArray(byrefalias32, _.ARGS.Ref(intindex, v143 => { intindex = v143; }), _.ARGS.Val("ReqNo")));
                        }
                        if (!loopConstraintsInitialized)
                            break;
                        var continueLoop = false;
                        _.HANDLEERROR(errOn, () => {
                            intindex = _.ADD(intindex, (Int16)1);
                            continueLoop = _.StrictLTE(intindex, loopEnd16);
                        });
                        if (!continueLoop)
                            break;
                    }
                }
            }

            // If there is at least one possible requirement match, add entry to list
            // (Otherwise, we can't do anything with the selection so don't bother with it)
            if (_.IF(_.GT(_.NullableNUM(_.CALL(this, _.CALL(this, objentry, _.ARGS.Val("PossReqNos")), "Count")), (Int16)0)))
            {
                object byrefalias33 = arrrequnitselections;
                _.CALL(this, byrefalias33, "Add", _.ARGS.Ref(objentry, v146 => { objentry = v146; }));
            }

            _.RELEASEERRORTRAPPINGTOKEN(errOn);
            return BookingUI_UnitSel_AddReqUnitSelection_retVal;
        }

        public object bookingui_unitsel_getmatchedrequnitselection(ref object arrrequnitoptions, ref object arrrequnitselections)
        {
            object BookingUI_UnitSel_GetMatchedReqUnitSelection_retVal = null;
            object lspermutations = null;
            object lstemp = null;
            object lspossreqnos = null;
            object intindex = null;
            object intindexsel = null;
            object intindexposs = null;
            object intindexperm = null;
            object intindexoption = null;
            object intscore = null;
            object intbestscore = null;
            object strbestpermutation = null;
            object arrmatches = null;
            object intunitkey = null;
            object lsunitkeys = null;
            object getmatchedrequnitselection = null;
            // Given list of requirement option objects and unit selection objects, try to match them up.

            // Ensure we've got values for both lists
            if (_.IF(_.OR(_.ISNULL(arrrequnitoptions), _.ISNULL(arrrequnitselections))))
            {
                getmatchedrequnitselection = VBScriptConstants.Nothing;
            }
            if (_.IF(_.OR(_.EQ(_.NullableNUM(_.CALL(this, arrrequnitoptions, "Count")), (Int16)0), _.EQ(_.NullableNUM(_.CALL(this, arrrequnitselections, "Count")), (Int16)0))))
            {
                getmatchedrequnitselection = VBScriptConstants.Nothing;
            }

            // First, create a list of ways in which the unit selections could be applied to the unit
            // options. We'll get out a list of strings which are comma-separated lists; the values
            // will relate the arrReqUnitSelections list indices to arrReqUnitOptions entries.
            //  eg. string "2,3,1"
            //      maps Selection 1 -> Option 2
            //           Selection 2 -> Option 3
            //           Selection 3 -> Option 1
            lspermutations = _.OBJ(_.CALL(this, _outer.page, "Functions", "GetNewObject", _.ARGS.Val("clsList")));
            var loopEnd17 = _.NUM(_.SUBT(_.CALL(this, arrrequnitselections, "Count"), (Int16)1));
            var loopStart17 = _.NUM((Int16)0, loopEnd17, (Int16)1);
            if (_.StrictLTE(loopStart17, loopEnd17))
            {
                for (intindexsel = loopStart17; _.StrictLTE(intindexsel, loopEnd17); intindexsel = _.ADD(intindexsel, (Int16)1))
                {
                    lspossreqnos = _.OBJ(_.CALL(this, _.CALL(this, arrrequnitselections, _.ARGS.Ref(intindexsel, v147 => { intindexsel = v147; })), _.ARGS.Val("PossReqNos")));
                    if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, lspermutations, "Count")), (Int16)0)))
                    {
                        // This is the first pass, so initialise the permutations list with
                        // the possible matches from this first ReqUnitSelection
                        var loopEnd18 = _.NUM(_.SUBT(_.CALL(this, lspossreqnos, "Count"), (Int16)1));
                        var loopStart18 = _.NUM((Int16)0, loopEnd18, (Int16)1);
                        if (_.StrictLTE(loopStart18, loopEnd18))
                        {
                            for (intindexposs = loopStart18; _.StrictLTE(intindexposs, loopEnd18); intindexposs = _.ADD(intindexposs, (Int16)1))
                            {
                                _.CALL(this, lspermutations, "Add", _.ARGS.RefIfArray(lspossreqnos, _.ARGS.Ref(intindexposs, v148 => { intindexposs = v148; })));
                            }
                        }
                    }
                    else
                    {
                        // We want to take our whatever permutation strings we have so far and expand
                        // them to include the possibilities for this ReqUnitSelection
                        // - Make a copy of lsPermutations thus far
                        lstemp = _.OBJ(_.CALL(this, _outer.page, "Functions", "GetNewObject", _.ARGS.Val("clsList")));
                        var loopEnd19 = _.NUM(_.SUBT(_.CALL(this, lspermutations, "Count"), (Int16)1));
                        var loopStart19 = _.NUM((Int16)0, loopEnd19, (Int16)1);
                        if (_.StrictLTE(loopStart19, loopEnd19))
                        {
                            for (intindexperm = loopStart19; _.StrictLTE(intindexperm, loopEnd19); intindexperm = _.ADD(intindexperm, (Int16)1))
                            {
                                _.CALL(this, lstemp, "Add", _.ARGS.RefIfArray(lspermutations, _.ARGS.Ref(intindexperm, v151 => { intindexperm = v151; })));
                            }
                        }
                        // - Clear out permutation list
                        _.CALL(this, lspermutations, "Clear");
                        // - Re-create new list using previous values with new combinations
                        var loopEnd20 = _.NUM(_.SUBT(_.CALL(this, lspossreqnos, "Count"), (Int16)1));
                        var loopStart20 = _.NUM((Int16)0, loopEnd20, (Int16)1);
                        if (_.StrictLTE(loopStart20, loopEnd20))
                        {
                            for (intindexposs = loopStart20; _.StrictLTE(intindexposs, loopEnd20); intindexposs = _.ADD(intindexposs, (Int16)1))
                            {
                                var loopEnd21 = _.NUM(_.SUBT(_.CALL(this, lstemp, "Count"), (Int16)1));
                                var loopStart21 = _.NUM((Int16)0, loopEnd21, (Int16)1);
                                if (_.StrictLTE(loopStart21, loopEnd21))
                                {
                                    for (intindexperm = loopStart21; _.StrictLTE(intindexperm, loopEnd21); intindexperm = _.ADD(intindexperm, (Int16)1))
                                    {
                                        _.CALL(this, lspermutations, "Add", _.ARGS.Val(_.CONCAT(_.CALL(this, lstemp, _.ARGS.Ref(intindexperm, v154 => { intindexperm = v154; })), ",", _.CALL(this, lspossreqnos, _.ARGS.Ref(intindexposs, v155 => { intindexposs = v155; })))));
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Now determine which arrangement matches the most selection / options pairs
            intbestscore = (Int16)(-1);
            var loopEnd22 = _.NUM(_.SUBT(_.CALL(this, lspermutations, "Count"), (Int16)1));
            var loopStart22 = _.NUM((Int16)0, loopEnd22, (Int16)1);
            if (_.StrictLTE(loopStart22, loopEnd22))
            {
                for (intindex = loopStart22; _.StrictLTE(intindex, loopEnd22); intindex = _.ADD(intindex, (Int16)1))
                {
                    intscore = _.VAL(_.CALL(this, _outer, "BookingUI_UnitSel_ScoreUnitSelPermutation", _.ARGS.RefIfArray(lspermutations, _.ARGS.Ref(intindex, v158 => { intindex = v158; }))));
                    if (_.IF(_.GT(intscore, intbestscore)))
                    {
                        intbestscore = _.VAL(intscore);
                        strbestpermutation = _.VAL(_.CALL(this, lspermutations, _.ARGS.Ref(intindex, v161 => { intindex = v161; })));
                    }
                }
            }

            // Finally, translate these matches into UnitKey values (or zero for unit
            // option which don't have a selection matched to them)
            // - Start off with a full-size list (matching size of arrReqUnitOptions) with
            //   with all zero values
            lsunitkeys = _.OBJ(_.CALL(this, _outer.page, "Functions", "GetNewObject", _.ARGS.Val("clsList")));
            var loopEnd23 = _.NUM(_.SUBT(_.CALL(this, arrrequnitoptions, "Count"), (Int16)1));
            var loopStart23 = _.NUM((Int16)0, loopEnd23, (Int16)1);
            if (_.StrictLTE(loopStart23, loopEnd23))
            {
                for (intindex = loopStart23; _.StrictLTE(intindex, loopEnd23); intindex = _.ADD(intindex, (Int16)1))
                {
                    _.CALL(this, lsunitkeys, "Add", _.ARGS.Val((Int16)0));
                }
            }

            // - Now push in the selection matches we have
            //    > Split best permutation back into integer values in arrMatches
            //    > The index of arrMatches will matches the index of arrReqUnitSelections
            //    > The value of arrMatches(n) will be the ReqNo it matches, which is the index
            //      of arrReqUnitOptions + 1 (andso also the index of lsUnitKeys + 1 since these
            //      two lists overlay)
            arrmatches = _.SPLIT(strbestpermutation, ",");
            var loopEnd24 = _.UBOUND(arrmatches);
            var loopStart24 = _.NUM((Int16)0, loopEnd24, (Int16)1);
            if (_.StrictLTE(loopStart24, loopEnd24))
            {
                for (intindexsel = loopStart24; _.StrictLTE(intindexsel, loopEnd24); intindexsel = _.ADD(intindexsel, (Int16)1))
                {
                    intindexoption = _.SUBT(_.CALL(this, arrmatches, _.ARGS.Ref(intindexsel, v162 => { intindexsel = v162; })), (Int16)1);
                    intunitkey = _.VAL(_.CALL(this, _.CALL(this, arrrequnitselections, _.ARGS.Ref(intindexsel, v163 => { intindexsel = v163; })), _.ARGS.Val("UnitKey")));
                    _.SET(_.VAL(intunitkey), this, lsunitkeys, null, _.ARGS.Ref(intindexoption, v165 => { intindexoption = v165; }));
                }
            }

            // Return matches!
            // There are the same number of values in lsUnitKeys as in arrReqUnitSelections, and
            // each lsUnitKeys(n) is the UnitKey for arrReqUnitSelections(n)
            BookingUI_UnitSel_GetMatchedReqUnitSelection_retVal = _.OBJ(lsunitkeys);

            return BookingUI_UnitSel_GetMatchedReqUnitSelection_retVal;
        }

        public object bookingui_unitsel_scoreunitselpermutation(ref object strpermutation)
        {
            object BookingUI_UnitSel_ScoreUnitSelPermutation_retVal = null;
            object intindex = null;
            object intscore = null;
            object arrvalues = null;
            object lsreqnos = null;
            // Determine a score for the Unit Selection / Option permutations calculated above.
            // Basically, give a score of one for each non-duplicated match.

            lsreqnos = _.OBJ(_.CALL(this, _outer.page, "Functions", "GetNewObject", _.ARGS.Val("clsList")));
            arrvalues = _.SPLIT(strpermutation, ",");
            intscore = (Int16)0;
            var loopEnd25 = _.UBOUND(arrvalues);
            var loopStart25 = _.NUM((Int16)0, loopEnd25, (Int16)1);
            if (_.StrictLTE(loopStart25, loopEnd25))
            {
                for (intindex = loopStart25; _.StrictLTE(intindex, loopEnd25); intindex = _.ADD(intindex, (Int16)1))
                {
                    if (_.IF(_.NOT(_.CALL(this, lsreqnos, "Contains", _.ARGS.RefIfArray(arrvalues, _.ARGS.Ref(intindex, v166 => { intindex = v166; }))))))
                    {
                        intscore = _.ADD(intscore, (Int16)1);
                        _.CALL(this, lsreqnos, "Add", _.ARGS.RefIfArray(arrvalues, _.ARGS.Ref(intindex, v169 => { intindex = v169; })));
                    }
                }
            }

            BookingUI_UnitSel_ScoreUnitSelPermutation_retVal = _.VAL(intscore);

            return BookingUI_UnitSel_ScoreUnitSelPermutation_retVal;
        }

        // ====================================================================================================
        // RENDER: Render options for accommodation products (only used with non-precise fuzzy stays)
        // ====================================================================================================
        // SUMMARY: summarise STAYS for this product which match booking criteria
        // [arsAvail]: ADO unit recordset from availability object
        // [adtStartNight]: date of first night of stay
        // [aiReqNumNights]: integer requested num nights
        public object bookingui_staysummary(ref object dtreqfirstnight, ref object ireqnights, ref object dtstayfirstnight, ref object istaynights, ref object po)
        {
            object BookingUI_StaySummary_retVal = null;

            // Render each stay result with link to further details
            // - 2009-08-10 DWR: Why do we not render this if "_stay" is in the querystring???
            if (_.IF(_.NOTEQ(_.NullableSTR(_.CALL(this, _outer.request, _.ARGS.Val("_stay"))), "")))
            {
                return BookingUI_StaySummary_retVal;
            }

            _.CALL(this, po, "Write", _.ARGS.Val("<div class=\"StayCandidateList\">"));
            _.CALL(this, po, "Write", _.ARGS.Val("<div class=\"StayCandidatesTtl\">"));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<p>", _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/flexiblesearchresults").Val("Flexible Search Results")), "</p>")));
            _.CALL(this, po, "Write", _.ARGS.Val("</div>"));
            if (_.IF(_.OR(_.NOTEQ(dtstayfirstnight, dtreqfirstnight), _.NOTEQ(ireqnights, istaynights))))
            {
                _.CALL(this, po, "Write", _.ARGS.Val("<div class=\"cell\">"));
                _.CALL(this, po, "Write", _.ARGS.Val("<div class=\"pnStayTtl\">"));
                object byrefalias34 = dtstayfirstnight, byrefalias35 = istaynights;
                try
                {
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer, "BookingUI_StayTtl", _.ARGS.Ref(byrefalias34, v172 => { byrefalias34 = v172; }).Ref(byrefalias35, v173 => { byrefalias35 = v173; }))));
                }
                finally { dtstayfirstnight = byrefalias34; istaynights = byrefalias35; }
                _.CALL(this, po, "Write", _.ARGS.Val("</div>"));
                object byrefalias36 = dtreqfirstnight, byrefalias37 = dtstayfirstnight, byrefalias38 = ireqnights, byrefalias39 = istaynights;
                try
                {
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer, "BookingUI_StayDiff", _.ARGS.Ref(byrefalias36, v174 => { byrefalias36 = v174; }).Ref(byrefalias37, v175 => { byrefalias37 = v175; }).Ref(byrefalias38, v176 => { byrefalias38 = v176; }).Ref(byrefalias39, v177 => { byrefalias39 = v177; }))));
                }
                finally { dtreqfirstnight = byrefalias36; dtstayfirstnight = byrefalias37; ireqnights = byrefalias38; istaynights = byrefalias39; }
                _.CALL(this, po, "Write", _.ARGS.Val("</div>"));
            }
            _.CALL(this, po, "Write", _.ARGS.Val("</div>"));

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
        public object bookingui_staydetails(object objavailentry, object istaynum, object adtstartnight, object aireqnights, object btelebooking, object strproductbookingwebifany, object streviivoidifany, object intproductkey, object strchannel, object bindicative, object bindicativevalid, object objallunits, object lsremoteunitselections, object po, object brendermaximumunitsavailable)
        {
            object BookingUI_StayDetails_retVal = null;
            object intindexunit = null;
            object objunit = null;
            object ilastreqmnt = null;
            object ithisreqmnt = null;
            object bgotopenreqcontainer = null;
            object sclassname = null;
            object bprecise = null;
            object iunitkey = null;
            object imaxrq = null;
            object iremoteunitkey = null;
            object bselected = null;
            object strnonbookableunits = null;
            object bhasbookableunits = null;
            object bhasnonbookableunits = null;

            // Ensure we've actually got some availability (we should if we've got here!)
            if (_.IF(_.EQ(_.NullableNUM(_.CALL(this, objavailentry, "Units", "Count")), (Int16)0)))
            {
                _.CALL(this, _outer.page, "PrintTraceWarning", _.ARGS.Val("BookingUI_StayDetails: No units in objAvailEntry"));
                return BookingUI_StayDetails_retVal;
            }

            // This method opens a new div - we'll need to close it later
            _.CALL(this, _outer, "BookingUI_RenderNewStay", _.ARGS.Ref(objavailentry, v178 => { objavailentry = v178; }).Ref(istaynum, v179 => { istaynum = v179; }).Ref(adtstartnight, v180 => { adtstartnight = v180; }).Ref(aireqnights, v181 => { aireqnights = v181; }).Ref(po, v182 => { po = v182; }));

            imaxrq = (Int16)0;
            ilastreqmnt = (Int16)0;
            iremoteunitkey = (Int16)0;
            bgotopenreqcontainer = false;
            bhasbookableunits = false;
            bhasnonbookableunits = false;
            var loopEnd26 = _.NUM(_.SUBT(_.CALL(this, objavailentry, "Units", "Count"), (Int16)1));
            var loopStart26 = _.NUM((Int16)0, loopEnd26, (Int16)1);
            if (_.StrictLTE(loopStart26, loopEnd26))
            {
                for (intindexunit = loopStart26; _.StrictLTE(intindexunit, loopEnd26); intindexunit = _.ADD(intindexunit, (Int16)1))
                {
                    objunit = _.OBJ(_.CALL(this, objavailentry, "Units", "GetItem", _.ARGS.Ref(intindexunit, v183 => { intindexunit = v183; })));

                    ithisreqmnt = _.VAL(_.CALL(this, objunit, "ReqNo"));
                    if (_.IF(_.GT(ithisreqmnt, imaxrq)))
                    {
                        // Moved on to next requirement, get key of pre-selected unit - iRemoteUnitKey
                        // will be zero if no selection has been passed in (applies to deep-linking)
                        imaxrq = _.VAL(ithisreqmnt);
                        iremoteunitkey = _.VAL(_.CALL(this, _outer, "BookingUI_GetPreSelectedUnitKey", _.ARGS.Ref(lsremoteunitselections, v184 => { lsremoteunitselections = v184; }).Ref(ithisreqmnt, v185 => { ithisreqmnt = v185; })));
                    }

                    // Check whether we're moving into a new requirement (if so, default to having
                    // the first unit appear selected) and render the "Room 1 - for 1 Guest"
                    // content
                    if (_.IF(_.NOTEQ(ithisreqmnt, ilastreqmnt)))
                    {

                        // If we've already got one of these containers open, close its tags
                        if (_.IF(bgotopenreqcontainer))
                        {
                            _.CALL(this, po, "Write", _.ARGS.Val("</div></div>"));
                        }
                        _.CALL(this, _outer, "BookingUI_RenderNewReq", _.ARGS.Ref(objunit, v186 => { objunit = v186; }).Ref(istaynum, v187 => { istaynum = v187; }).Ref(ithisreqmnt, v188 => { ithisreqmnt = v188; }).Val(_.NOT(_.CALL(this, objavailentry, "IsLocal"))).Ref(po, v189 => { po = v189; }));
                        bgotopenreqcontainer = true;

                        bselected = true;
                        ilastreqmnt = _.VAL(ithisreqmnt);
                    }
                    else
                    {
                        bselected = false;
                    }

                    // .. however, if there was a pre-selected unit key passed in, this should override which
                    // unit appears selected (this only applies when iRemoteUnitKey is not zero, meaning that
                    // a unit selection exists - note: eviivo units always appear with unit key zero)
                    iunitkey = _.VAL(_.CALL(this, objunit, "UnitKey"));
                    if (_.IF(_.NOTEQ(_.NullableNUM(iremoteunitkey), (Int16)0)))
                    {
                        bselected = _.VAL(_.EQ(iunitkey, iremoteunitkey));
                    }

                    // build up a list of invalid indicative or telephone booking
                    // units, this is used later by javascript when we have a mixture of allocated and indicative
                    // availability
                    if (_.IF(_.OR(_.AND(_.CALL(this, objunit, "Indicative"), _.NOT(bindicativevalid)), btelebooking)))
                    {
                        bhasnonbookableunits = true;
                        if (_.IF(_.GT(_.NullableNUM(_.LEN(strnonbookableunits)), (Int16)0)))
                        {
                            strnonbookableunits = _.CONCAT(strnonbookableunits, ",");
                        }

                        //MJ - 	the stay num is no longer part of this data, it is part of each array's name
                        //		look at TB's other changes to see the reasoning behind this
                        strnonbookableunits = _.CONCAT(strnonbookableunits, iunitkey);
                        _.CALL(this, _outer.page, "PrintTrace", _.ARGS.Val(_.CONCAT("strNonBookableUnits", strnonbookableunits)));
                    }
                    else
                    {
                        bhasbookableunits = true;
                    }

                    // 2009-09-30 DWR: The AvailClassName was previously generated by considering the indicative
                    // state of the whole stay - this was causing all units to be rendered as indicative if any
                    // one of them was, now we take the indicative state from each unit (but keep the indicative
                    // "validity" from the whole stay, where required)
                    _.CALL(this, _outer, "BookingUI_RenderUnit", _.ARGS.Ref(istaynum, v190 => { istaynum = v190; }).Ref(ithisreqmnt, v191 => { ithisreqmnt = v191; }).Ref(bselected, v192 => { bselected = v192; }).Ref(objavailentry, v193 => { objavailentry = v193; }).Ref(objunit, v194 => { objunit = v194; }).Ref(objallunits, v195 => { objallunits = v195; }).Val(_.CALL(this, _outer, "BookingUI_AvailClassName", _.ARGS.Val(_.CALL(this, objunit, "Indicative")).Ref(bindicativevalid, v196 => { bindicativevalid = v196; }).Ref(btelebooking, v197 => { btelebooking = v197; }))).Ref(po, v198 => { po = v198; }).Ref(brendermaximumunitsavailable, v199 => { brendermaximumunitsavailable = v199; }));

                }
            }

            // Ensure any open req container (eg. "Room 1 - for 1 Guest" section) is closed
            if (_.IF(bgotopenreqcontainer))
            {
                _.CALL(this, po, "Write", _.ARGS.Val("</div></div>"));
                bgotopenreqcontainer = false;
            }

            // Close the BookingUI_RenderNewStay containing div
            _.CALL(this, po, "Write", _.ARGS.Val("</div>"));

            // Wrap these hidden inputs in a div for html validity
            _.CALL(this, po, "Write", _.ARGS.Val("<div>"));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"_nStays\" value=\"", istaynum, "\" />")));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"_nReqs\" value=\"", imaxrq, "\" />")));
            if (_.IF(_.NOT(_.CALL(this, objavailentry, "IsLocal"))))
            {
                _.CALL(this, po, "Write", _.ARGS.Val("<input type=\"hidden\" name=\"IsEviivoBooking\" value=\"yes\" />"));
                if (_.IF(_outer.isexternalbooking))
                {
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"eviivoconf\" value=\"", _.CLNG(_.CONCAT("0", _.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("Integration_Eviivo_ConfigSet")))), "\" />")));
                }
            }
            _.CALL(this, po, "Write", _.ARGS.Val("</div>"));

            // 2014-06-25 DWR: For sites that use the legacy "eviivo external" booking integration (meaning sites where VB Polling is not enabled - the new implementation
            // results in Eviivo results being reported as Polling results and the user being sent through the Polling Exit with a fully-populated deep link), the Book
            // button should not be shown here. The Unit Selection should never be shown in this case, to be honest, since Book buttons should go straight to the Product's
            // Booking Website and not enter the site's availability process. However, if there are sites that show inline Unit Selection (inline with the Product List)
            // then the Unit data may be useful. If we were wanted to render Book buttons here (to the external site) then logic would have to be duplicated from the
            // Product List or Detail Control, which would be better avoided. A much better solution is to enable VB Polling and avoid this legacy mechanism entirely.
            // Note: We could potentially render the button for Local Avail and not for Eviivo but I think that that's more confusing than helpful, particularly since
            // it's inconsistent with the Product List / Detail implementation (which bases its decision upon whether the Product has an Eviivo Id).
            if (_.IF(_.AND(_.AND(_.NOT(_.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("Booking_EnablePolling"))), _.NOTEQ(_.NullableSTR(_.TRIM(_.CONCAT("", streviivoidifany))), "")), _.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("Integration_Eviivo_ExtBooking_Enable")))))
            {
                _.CALL(this, _outer.page, "PrintTraceWarning", _.ARGS.Val("Not rendering any Book buttons for Unit Selection since the legacy Eviivo External Booking configuration is enabled (the recommended alternative is to use the deep-link-supporting Eviivo External Booking configuration, this may be done by enabling VB Polling)"));
                return BookingUI_StayDetails_retVal;
            }

            // 2014-03-14 DWR: New functionality "Availability Searches with offsite Booking Web Booking" allows for Products to be on the Telephone Booking Channel
            // and have their availability queried but to show a Booking button that goes to the Product's Booking Website (if one is specified), rather than
            // showing a "this can not be booked online, please call.." message (this means that the avail criteria have to be re-entered on the target
            // website, but that is understood and how it works - see FogBugz 10367). I've tried to make the markup for this button reminiscent of
            // that in Product List and Detail to try to make any additional styling requirements as low as possible.
            strproductbookingwebifany = _.VAL(_.TRIM(_.CONCAT("", strproductbookingwebifany)));
            if (_.IF(_.AND(_.AND(_.AND(btelebooking, _.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("Booking_EnableByPhone"))), _.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("Booking_AllowOffSiteTelephoneBookings"))), _.NOTEQ(_.NullableSTR(strproductbookingwebifany), ""))))
            {
                _.CALL(this, _outer.page, "PrintTrace", _.ARGS.Val("Since this is a Telephone Booking Product with a Booking Website and the 'Allow Offsite Booking Web Booking for Telephone Bookings' parameter is enabled, a button to the Booking Website is being rendered"));
                _.CALL(this, po, "Write", _.ARGS.Val("<div class=\"pnStayButtons\">"));
                _.CALL(this, po, "Write", _.ARGS.Val("<p class=\"bookonline\">"));
                _.CALL(this, po, "Write", _.ARGS.Val("<a href=\""));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer.server, "HtmlEncode", _.ARGS.Ref(strproductbookingwebifany, v200 => { strproductbookingwebifany = v200; }))));
                _.CALL(this, po, "Write", _.ARGS.Val("\""));
                if (_.IF(_.OR(_.CALL(this, _outer.page, "IsPartialRender"), _.EQ(_.NullableSTR(_.CALL(this, _outer.request, _.ARGS.Val("PartialRenderType"))), "html"))))
                {
                    // If in Partial Render then set target="_blank" instead of rel="external" (we only do the latter for strict adherence to standards and then
                    // use javascript to transform after rendering - when requesting additional content through javascript this transformation won't be performed
                    // so we'll need to generate it direct)
                    // 2014-06-12 DWR: The partial render requests for this data are commonly made as "html" meaning that Page.IsPartialRender will be false
                    // (the logic being that Controls should render entirely as standard when in html partial render mode) so I've added an additional check
                    // for the a "PartialRenderType" value of "html" to ensure that the new-window logic is maintained correctly.
                    _.CALL(this, po, "Write", _.ARGS.Val(" target=\"_blank\""));
                }
                else
                {
                    _.CALL(this, po, "Write", _.ARGS.Val(" rel=\"external\""));
                }
                _.CALL(this, po, "Write", _.ARGS.Val(" class=\"ProvClickCustom\" name=\"PROBWEBREF|"));
                // This is the "Provider Booking Website Referral" statistic, as required by the SharePoint document for FogBugz 10367
                _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer.server, "HtmlEncode", _.ARGS.Ref(strchannel, v201 => { strchannel = v201; }))));
                _.CALL(this, po, "Write", _.ARGS.Val("|"));
                _.CALL(this, po, "Write", _.ARGS.Ref(intproductkey, v202 => { intproductkey = v202; }));
                _.CALL(this, po, "Write", _.ARGS.Val("\""));
                _.CALL(this, po, "Write", _.ARGS.Val(">"));
                _.CALL(this, po, "Write", _.ARGS.Val("<img src=\""));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer.page, "ImageResource", _.ARGS.Val("bookonline/btn/book").Val(_.CONCAT(_.CALL(this, _outer.context, "ImageDir"), "booking/book.gif")))));
                _.CALL(this, po, "Write", _.ARGS.Val("\" alt=\""));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/btn/book").Val("Book"))));
                _.CALL(this, po, "Write", _.ARGS.Val(" ("));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("productdetail/bookonline/opensinanewwindow").Val("opens in a new window"))));
                _.CALL(this, po, "Write", _.ARGS.Val(")\" "));
                _.CALL(this, po, "Write", _.ARGS.Val("/>"));
                _.CALL(this, po, "Write", _.ARGS.Val("</a>"));
                _.CALL(this, po, "Write", _.ARGS.Val("</p>"));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("</div>", VBScriptConstants.vbCrLf)));
                return BookingUI_StayDetails_retVal;
            }

            // 2014-03-13 DWR: If there is at least one bookable unit then display the Book button and rely on JavaScript to show/hide it if selections are made that
            // can not be completed online. But if there are NO bookable units (eg. a Telephone Booking Product or all of the Units are Indicative where the timeout
            // period has passed) then there's no point even rendering the button.
            if (_.IF(bhasbookableunits))
            {
                _.CALL(this, _outer, "BookingUI_RenderButtons", _.ARGS.Ref(istaynum, v203 => { istaynum = v203; }).Ref(po, v204 => { po = v204; }).Val(_.CALL(this, objavailentry, "IsExternal")));
            }

            // if we have an invalid indicative unit or telephone unit then
            // render this message - let the js do the rest
            if (_.IF(bhasnonbookableunits))
            {

                // 2010-07-09 PW: RIP Gary
                // This is the array formerly known as garyTeleBookUnitKeys
                // it is used for switching between the online book button if the unit is bookable
                // or rendering the relevant warning message if it isn't
                // 2010-10-21 TB: augmenting Gary with stay key. This is to allow for multiple stays
                // in which this JS is executed on a per stay basis via a partial render.
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<script type=\"text/javascript\">", VBScriptConstants.vbCrLf)));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT(" var aryNonBookableUnits_", istaynum, " = [", strnonbookableunits, "]; ", VBScriptConstants.vbCrLf)));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT(" var iTotalNonBookableUnits = ", ithisreqmnt, ";", VBScriptConstants.vbCrLf)));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("</", "script>")));

                // Render relevant offline booking message
                _.CALL(this, po, "Write", _.ARGS.Val("<div id=\"pnTeleBook_PromptCall\">"));
                if (_.IF(_.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("Booking_EnableByPhone"))))
                {
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<p>", _.REPLACE(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/telebook/prompt").Val("One or more of the units you have selected must be booked via telephone. Please ring #bookingtelephone# to continue this booking.")), "#bookingtelephone#", _.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("Booking_TelephoneNumber"))), "</p>")));
                }
                else
                {
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<p>", _.REPLACE(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/indtelebook/prompt").Val("Although available, some of the units you have selected cannot be booked online. Alternatively, select different units with online booking only.")), "#bookingtelephone#", _.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("Booking_TelephoneNumber"))), "</p>")));
                }
                _.CALL(this, po, "Write", _.ARGS.Val("</div>"));
            }

            return BookingUI_StayDetails_retVal;
        }

        public object bookingui_getpreselectedunitkey(object lsremoteunitselections, object ireqno)
        {
            object BookingUI_GetPreSelectedUnitKey_retVal = null;
            // remote referrals (eg. VB integration) will include a UNIT_KEY choice and CHILD_COUNT.
            // eg. Request vars formatted as 'URslt[REQ_NUMBER]=[UNIT_KEY]-[NUM_ADULT]-[NUM-CHILD]'

            // If not got here from a remote referral (eg. VB integration), the lsRemoveUnitSelections will be Nothing
            if (_.IF(_.NOT(_.IS(lsremoteunitselections, VBScriptConstants.Nothing))))
            {
                // Get unit selection passed in (may be zero if invalid request was made)
                // Note: lsRemoteUnitSelections has zero-based index, iReqNo is one-based
                if (_.IF(_.AND(_.GTE(_.NullableNUM(ireqno), (Int16)1), _.LTE(ireqno, _.CALL(this, lsremoteunitselections, "Count")))))
                {
                    BookingUI_GetPreSelectedUnitKey_retVal = _.VAL(_.CALL(this, lsremoteunitselections, _.ARGS.Val(_.SUBT(ireqno, (Int16)1))));
                    return BookingUI_GetPreSelectedUnitKey_retVal;
                }
            }

            BookingUI_GetPreSelectedUnitKey_retVal = (Int16)0;
            return BookingUI_GetPreSelectedUnitKey_retVal;
        }

        // SUMMARY: for VB Polling - we want to render a supplier name and icon above each set of unit options
        public object bookingui_staydetails_pollingheader(object objavailentry, object po, object strsupplierlogo, object strsuppliername)
        {
            object BookingUI_StayDetails_PollingHeader_retVal = null;

            // Render header content (icon, if specified) and supplier name
            // 2008-12-18 DWR: Add a style to indicate whether supplier is Local, FrontDesk or External (this will
            // allow a custom logo to be used for Local or FrontDesk, for example)
            _.CALL(this, po, "Write", _.ARGS.Val("<div class=\"StayCandidateItemHeader "));
            if (_.IF(_.CALL(this, objavailentry, "IsLocal")))
            {
                _.CALL(this, po, "Write", _.ARGS.Val(" AvailLocal"));
            }
            else if (_.IF(_.CALL(this, objavailentry, "IsRemote")))
            {
                _.CALL(this, po, "Write", _.ARGS.Val(" AvailFrontDesk"));
            }
            else
            {
                _.CALL(this, po, "Write", _.ARGS.Val(" AvailExternal"));
            }
            _.CALL(this, po, "Write", _.ARGS.Val("\">"));
            if (_.IF(_.NOTEQ(_.NullableSTR(strsupplierlogo), "")))
            {
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<img src=\"", strsupplierlogo, "\" alt=\"", strsuppliername, "\" />")));
            }
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<h2>", strsuppliername, "</h2>")));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("</div>", VBScriptConstants.vbCrLf)));
            return BookingUI_StayDetails_PollingHeader_retVal;
        }

        //tries to get a supplier logo for us
        public object getsupplierlogo(ref object strproductestateid)
        {
            object GetSupplierLogo_retVal = null;
            object strsupplierlogo = null;
            strsupplierlogo = "";
            if (_.IF(_outer.isexternalbooking))
            {
                strsupplierlogo = _.VAL(_.CALL(this, _outer.page, "ImageResource", _.ARGS.Val(_.CONCAT("bookonline/unitselection/polling/localsupplier/estate_", strproductestateid, "/logo")).Val("")));
                if (_.IF(_.EQ(_.NullableSTR(strsupplierlogo), "")))
                {
                    strsupplierlogo = _.VAL(_.CALL(this, _outer.page, "Resource", _.ARGS.Val(_.CONCAT("bookonline/unitselection/polling/localsupplier/estate_", strproductestateid, "/logo")).Val("")));
                    if (_.IF(_.NOTEQ(_.NullableSTR(strsupplierlogo), "")))
                    {
                        _.CALL(this, _outer.page, "PrintTraceWarning", _.ARGS.Val("Loaded estate scoped supplier logo from a deprecated location - please move it to the image resources language file"));
                    }
                }
            }
            if (_.IF(_.EQ(_.NullableSTR(strsupplierlogo), "")))
            {
                strsupplierlogo = _.VAL(_.CALL(this, _outer.page, "ImageResource", _.ARGS.Val("bookonline/unitselection/polling/localsupplier/logo").Val("")));
                if (_.IF(_.EQ(_.NullableSTR(strsupplierlogo), "")))
                {
                    strsupplierlogo = _.VAL(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/polling/localsupplier/logo").Val("")));
                    if (_.IF(_.NOTEQ(_.NullableSTR(strsupplierlogo), "")))
                    {
                        _.CALL(this, _outer.page, "PrintTraceWarning", _.ARGS.Val("Loaded estate scoped supplier logo from a deprecated location - please move it to the image resources language file"));
                    }
                }
            }
            GetSupplierLogo_retVal = _.VAL(strsupplierlogo);
            return GetSupplierLogo_retVal;
        }

        // SUMMARY: return URL which browsers without Javascript can use to navigate stay candidates page
        // [aiStay]: integer stay number. 1 = 1st stay, 2 = 2nd stay. Zero produces back URL to stay candidates page
        // <retval>: string URL for hyperlink
        public object bookingui_staydetailsurl(ref object aistay)
        {
            object BookingUI_StayDetailsUrl_retVal = null;
            object surl = null;
            object sstay = null;
            object ipos = null;
            object sright = null;
            object iright = null; /* Undeclared in source */

            // get current URL. Prepare [stay] variable to be appended to URL
            surl = _.VAL(_.CALL(this, _outer.request, "ServerVariables", _.ARGS.Val("HTTP_X_REWRITE_URL")));
            if (_.IF(_.GT(_.NullableNUM(aistay), (Int16)0)))
            {
                sstay = _.CONCAT("&_stay=", aistay);
            }
            else
            {
                sstay = "";
            }

            // does URL already have stay variable? if so, remove it and return new URL
            ipos = _.VAL(_.INSTR(surl, "&_stay="));
            if (_.IF(_.GT(_.NullableNUM(ipos), (Int16)0)))
            {
                sright = _.VAL(_.MID(surl, _.ADD(ipos, (Int16)7)));
                iright = _.VAL(_.INSTR(sright, "&"));
                surl = _.VAL(_.LEFT(surl, _.SUBT(ipos, (Int16)1)));
                if (_.IF(_.GT(_.NullableNUM(iright), (Int16)0)))
                {
                    surl = _.CONCAT(surl, _.MID(sright, iright));
                }
            }
            BookingUI_StayDetailsUrl_retVal = _.CONCAT(surl, sstay);
            return BookingUI_StayDetailsUrl_retVal;
        }

        // SUMMARY: render new stay UI - WARNING: this doesn't close all of the elements it opens!
        // [objAvailEntry]: avail data for a single stay
        // [aiStayNum]: integer stay index (1-based)
        // [adtStartNight]: date requested start night
        // [aiReqNights]: integer requested num nights
        public object bookingui_rendernewstay(object objavailentry, object aistaynum, object adtstartnight, object aireqnights, object po)
        {
            object BookingUI_RenderNewStay_retVal = null;
            object spostfix = null;
            object bprecise = null;
            object bexactmatch = null;

            // Render slightly differently if got a precise match
            // - Also render differently when VB Polling enabled, since we have to render
            //   more of these sections than otherwise
            bexactmatch = _.VAL(_.AND(_.EQ(_.CALL(this, objavailentry, "StartDate"), adtstartnight), _.EQ(_.CALL(this, objavailentry, "Nights"), aireqnights)));
            if (_.IF(_.OR(bexactmatch, _outer.isvbpollingenabled)))
            {
                bprecise = true;
                spostfix = "1";
            }
            else if (_.IF(_.EQ(_.CLNG(_.CONCAT("0", _.CALL(this, _outer.request, _.ARGS.Val("_stay")))), aistaynum)))
            {
                spostfix = "1";
            }
            else
            {
                spostfix = "";
            }

            // If not exact match then render a warning as well as the date difference later
            if (_.IF(_.NOT(bexactmatch)))
            {
                _.CALL(this, _outer, "RenderNotRequiredDateWarning", _.ARGS.Ref(po, v205 => { po = v205; }));
            }

            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<div class=\"StayCandidateItem", spostfix, "\">", VBScriptConstants.vbCrLf)));

            if (_.IF(_.NOT(bexactmatch)))
            {
                _.CALL(this, po, "Write", _.ARGS.Val("<div class=\"pnStayTtl\">"));
                _.CALL(this, po, "Write", _.ARGS.Val("<p>"));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer, "BookingUI_StayTtl", _.ARGS.Val(_.CALL(this, objavailentry, "StartDate")).Val(_.CALL(this, objavailentry, "Nights")))));
                _.CALL(this, po, "Write", _.ARGS.Val("</p>"));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("</div>", VBScriptConstants.vbCrLf)));
                if (_.IF(_.NOT(_outer.brenderascalendar)))
                {
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer, "BookingUI_StayDiff", _.ARGS.Ref(adtstartnight, v206 => { adtstartnight = v206; }).Val(_.CALL(this, objavailentry, "StartDate")).Ref(aireqnights, v207 => { aireqnights = v207; }).Val(_.CALL(this, objavailentry, "Nights")))));
                }
            }
            return BookingUI_RenderNewStay_retVal;
        }

        // SUMMARY: return title for this stay candidate
        // [aiNights]: integer number nights for this stay
        // [adtFirstNight]: date of first night
        // [adtLastNight]: date of last night
        // <retval>: string stay title
        public object bookingui_stayttl(object adtfirstnight, object ainights)
        {
            object BookingUI_StayTtl_retVal = null;
            if (_.IF(_.EQ(_.NullableNUM(ainights), (Int16)1)))
            {
                BookingUI_StayTtl_retVal = _.CONCAT(ainights, _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/nightstart").Val(" night, start ")), _.CALL(this, _outer.page, "Functions", "Dates", "ShortDate", _.ARGS.Ref(adtfirstnight, v208 => { adtfirstnight = v208; })));
                return BookingUI_StayTtl_retVal;
            }

            BookingUI_StayTtl_retVal = _.CONCAT(ainights, _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/nightsfrom").Val(" nights, from ")), _.CALL(this, _outer.page, "Functions", "Dates", "ShortDate", _.ARGS.Ref(adtfirstnight, v210 => { adtfirstnight = v210; })), _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/to").Val(" to ")), _.CALL(this, _outer.page, "Functions", "Dates", "Shortdate", _.ARGS.Val(_.DATEADD("d", ainights, adtfirstnight))));
            return BookingUI_StayTtl_retVal;
        }

        // SUMMARY: describe difference between THIS DATE and REQUESTED stay date
        // [adtReqDate]: date of REQUESTED first night of stay
        // [adtThisDate]: date of RESULTANT first night of stay
        // [aiReqNights]: integer requested num nights
        // [aiNights]: integer result num nights
        public object bookingui_staydiff(object adtreqdate, object adtthisdate, object aireqnights, object airesultnights)
        {
            object BookingUI_StayDiff_retVal = null;
            object idatediff = null;
            object idurdiff = null;

            idatediff = _.VAL(_.DATEDIFF("d", adtreqdate, adtthisdate));
            idurdiff = _.SUBT(airesultnights, aireqnights);
            BookingUI_StayDiff_retVal = _.CONCAT("<div class=\"pnStayDiff\">", _.CALL(this, _outer.page, "Functions", "Booking", "Booking_MatchQual", _.ARGS.Val((Int16)0).Ref(idatediff, v212 => { idatediff = v212; }).Ref(idurdiff, v213 => { idurdiff = v213; }).Ref(aireqnights, v214 => { aireqnights = v214; }).Val((Int16)2)), "</div>", VBScriptConstants.vbCrLf);
            return BookingUI_StayDiff_retVal;
        }

        // SUMMARY: render new requirement UI - WARNING: this doesn't close all of the elements it opens!
        // [arsAvail]: ADO unit recordset from availability object
        // [aiStayNum]: integer stay index
        // [aiThisReqmnt]: integer requirement number (from recordset)
        public object bookingui_rendernewreq(object objunit, object aistaynum, object aithisreqmnt, object abremote, object po)
        {
            object BookingUI_RenderNewReq_retVal = null;
            object isz = null;
            object sreqmntset = null;
            object sremoterqmnt = null;
            object ichild = null;
            object iremotenumchild = null;
            object sselected = null;
            object arychildages = null;
            object ichildageindex = null;

            isz = _.VAL(_.CALL(this, objunit, "ReqSize"));

            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<div class=\"pnStayReqmnt\">", VBScriptConstants.vbCrLf)));
            _.CALL(this, po, "Write", _.ARGS.Val("<div class=\"pnStayReqmntTtl\">"));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/room").Val("Room"))));
            _.CALL(this, po, "Write", _.ARGS.Val(" "));
            _.CALL(this, po, "Write", _.ARGS.Ref(aithisreqmnt, v218 => { aithisreqmnt = v218; }));
            _.CALL(this, po, "Write", _.ARGS.Val(" - "));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/for").Val("for"))));
            _.CALL(this, po, "Write", _.ARGS.Val(" "));
            _.CALL(this, po, "Write", _.ARGS.Ref(isz, v219 => { isz = v219; }));
            _.CALL(this, po, "Write", _.ARGS.Val(" "));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/guest(s)").Val("guest(s)"))));

            //#MJ -	We can only render our room requirement data based upon the recieved dat, not the requirement we passed in, as it may have been fulfilled in a different order
            //2012-03-29 NP: Here we render the requirements that are linked to the unit stay details in the response from the Avail Component
            // we do NOT want to render the original request against each unit that is rendered because they may not order up
            // Example: Request roomReq_1 = 2; roomReq_2 = 1; Response may come back in a different order
            // i.e. unit_1 with ReqSize = 1, unit_2 with ReqSize = 2 so roomReq_1 = 1, roomReq= 2; they end up swapped around
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"roomReq_", aithisreqmnt, "\" value=\"", isz, "\" />")));

            //#MJ - need to check with Rich if we want to indicate who's going into what room
            if (_.IF(_.AND(_.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("Booking_ChildPricing")), _.GT(_.NullableNUM(_.CALL(this, objunit, "ChildrenRequirement")), (Int16)0))))
            {
                _.CALL(this, po, "Write", _.ARGS.Val(" - ("));
                _.CALL(this, po, "Write", _.ARGS.Val("<span class=\"ReqmntDetails\">"));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("adults").Val("Adults"))));
                _.CALL(this, po, "Write", _.ARGS.Val(": "));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, objunit, "AdultsRequirement")));
                _.CALL(this, po, "Write", _.ARGS.Val(" "));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("children").Val("Children"))));
                _.CALL(this, po, "Write", _.ARGS.Val(": "));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, objunit, "ChildrenRequirement")));
                _.CALL(this, po, "Write", _.ARGS.Val(") "));
                _.CALL(this, po, "Write", _.ARGS.Val("</span>"));
                // NP 2012-03-01: Child pricing requirements were not previously being posted to the checkout
                // Adult & Child Requirement amount is needed by the RequirementSummary control and the child ages are
                // needed by the checkout for creating the correct requirement record with the relevant discount values
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"roomReq_", aithisreqmnt, "_adults\" value=\"", _.CALL(this, objunit, "AdultsRequirement"), "\" />")));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"roomReq_", aithisreqmnt, "_children\" value=\"", _.CALL(this, objunit, "ChildrenRequirement"), "\" />")));

                // ChildrenAges is a comma separated list of ages or "", Split will give an empty array if this property is ever Empty
                arychildages = _.SPLIT(_.CALL(this, objunit, "ChildrenAges"), ",");
                var loopEnd27 = _.UBOUND(arychildages);
                var loopStart27 = _.NUM((Int16)0, loopEnd27, (Int16)1);
                if (_.StrictLTE(loopStart27, loopEnd27))
                {
                    for (ichildageindex = loopStart27; _.StrictLTE(ichildageindex, loopEnd27); ichildageindex = _.ADD(ichildageindex, (Int16)1))
                    {
                        _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"hidden\" name=\"roomReq_", aithisreqmnt, "_children_childage", ichildageindex, "\" value=\"", _.CALL(this, arychildages, _.ARGS.Ref(ichildageindex, v220 => { ichildageindex = v220; })), "\" />")));
                    }
                }

            }

            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("</div>", VBScriptConstants.vbCrLf)));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<div class=\"pnStayReqmntRslts\">", VBScriptConstants.vbCrLf)));

            return BookingUI_RenderNewReq_retVal;
        }

        // SUMMARY: render unit option HTML
        // [aiStayNum]: integer stay index
        // [aiThisReqmnt]: integer requirement index
        // [aiUnitKey]: integer unit key
        // [bSelected]: should the current unit appear selected
        // [arsAvail]: ADO availability recordset
        // [asAvailClassName]: string avail class name
        public object bookingui_renderunit(object aistaynum, object aithisreqmnt, object bselected, object objavailentry, object objunit, object objallunits, object asavailclassname, object po, object brendermaximumunitsavailable)
        {
            object BookingUI_RenderUnit_retVal = null;
            object munitstaytotal = null;
            object inumnights = null;
            object munitpernight = null;
            object inumpeople = null;
            object idaysbreakfast = null;
            object bperperson = null;
            object mpersonpernight = null;
            object striptid = null;
            object munitstaytotalpayablebasedonguideprice = null;
            object bdiscountapplied = null;
            object iadults = null;
            object ichildren = null;
            object imaxunitsavailable = null;
            object unitcostperperson = null; /* Undeclared in source */

            munitstaytotal = _.VAL(_.CALL(this, objunit, "StayTotalPayable"));
            munitstaytotalpayablebasedonguideprice = _.VAL(_.CALL(this, objunit, "StayTotalPayableBasedOnGuidePrice"));
            inumnights = _.VAL(_.CALL(this, objavailentry, "Nights"));
            munitpernight = _.DIV(munitstaytotal, inumnights);
            bperperson = _.VAL(_.CALL(this, objunit, "Perperson"));
            inumpeople = _.VAL(_.CALL(this, objunit, "ReqSize"));

            idaysbreakfast = _.VAL(_.CALL(this, objunit, "DaysBreakfast"));
            bdiscountapplied = _.VAL(_.CALL(this, objunit, "IncludesChildDiscount"));

            imaxunitsavailable = _.VAL(_.CALL(this, objunit, "MaximumQuantityAvailable"));

            // We need an id so we can set the label's "for" attribute, but if VB Polling is enabled,
            // we might end up with id duplication - so in that case we append a random suffix
            striptid = _.CONCAT("unit_", aistaynum, "_", aithisreqmnt, "_", _.CALL(this, objunit, "UnitKey"));
            if (_.IF(_outer.isvbpollingenabled))
            {
                striptid = _.CONCAT(striptid, "_", _.INT(_.MULT(_.RND(), 100000)));
            }

            _.CALL(this, po, "Write", _.ARGS.Val("<div class=\"pnUnitOption\">"));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"radio\" name=\"unit_", aistaynum, "_", aithisreqmnt, "\" ", "id=\"", striptid, "\" ")));
            if (_.IF(_.NOT(_outer.isvbpollingenabled)))
            {
                // Not sure this onclick is even required without VB Polling.. (?)
                _.CALL(this, po, "Write", _.ARGS.Val("onclick=\"BookingUI_UnitSelect(this);\" "));
            }
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("value=\"", _.CALL(this, objunit, "UnitKey"), "\" ")));
            if (_.IF(bselected))
            {
                _.CALL(this, po, "Write", _.ARGS.Val("checked=\"checked\" "));
            }
            _.CALL(this, po, "Write", _.ARGS.Val("/>"));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<label for=\"", striptid, "\"> ")));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT(_.CALL(this, objunit, "UnitName"), " - ", _.CALL(this, _outer, "BookingUI_NicePrice", _.ARGS.Ref(munitstaytotal, v222 => { munitstaytotal = v222; })), " ", asavailclassname)));

            //if we have child pricing discount applied show the icon
            if (_.IF(bdiscountapplied))
            {
                _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer, "BookingUI_AvailClassIcon", _.ARGS.Val("DISCOUNT"))));
            }

            _.CALL(this, po, "Write", _.ARGS.Val("</label>"));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("</div>", VBScriptConstants.vbCrLf)));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<div class=\"pnPriceBase\">", VBScriptConstants.vbCrLf)));

            //#MJ 29/04/2010 -	decision made not to show the price basis as the per person figure was always a guestimate, child pricing messes with the price so per person doesn't apply
            //					also we now always deal with total stay prices
            //				If bPerPerson Then
            //					mPersonPerNight = mUnitPerNight/iNumPeople
            //					pO.Write BookingUI_NicePrice(mPersonPerNight) & " " & Page.Resource("bookonline/unitselection/perpersonpernight", "per person per night") & ". "
            //				Else
            //					pO.Write BookingUI_NicePrice(mUnitPerNight) & " " & Page.Resource("bookonline/unitselection/perroomunitpernight", "per room/unit per night") & ". "
            //				End If

            if (_.IF(_.EQ(idaysbreakfast, inumnights)))
            {
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/breakfastincluded").Val("Breakfast included")), ". ")));
            }
            else if (_.IF(_.GT(_.NullableNUM(idaysbreakfast), (Int16)0)))
            {
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/breakfastincludedon").Val("Breakfast included on ")), idaysbreakfast, " ", _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/day(s)").Val("day(s)")), ". ")));
            }

            if (_.IF(_.LT(inumpeople, _.CALL(this, objunit, "MinOcc"))))
            {
                if (_.IF(bperperson))
                {
                    _.CALL(this, po, "Write", _.ARGS.Val("<br />"));
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/priceperpersonincludes").Val("Price Per Person includes")), " ")));
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CALL(this, _outer, "BookingUI_NicePrice", _.ARGS.Val(_.SUBT(mpersonpernight, _.DIV(unitcostperperson, inumnights))))));
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/minimumoccupancysupplement").Val(" minimum occupancy supplement")), ". ")));
                }
                else
                {
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/minoccupancyof").Val("Min. occupancy of")), " ", _.CALL(this, objunit, "MinOcc"), ". ")));
                }
            }
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<div class=\"pnLinkedUnit\">", _.CALL(this, _outer, "BookingUI_LinkedUnitDesc", _.ARGS.Ref(objunit, v224 => { objunit = v224; }).Ref(objallunits, v225 => { objallunits = v225; })), "</div>")));

            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("</div>", VBScriptConstants.vbCrLf)));

            if (_.IF(_.NOT(_.CALL(this, objavailentry, "IsLocal"))))
            {
                _.CALL(this, po, "Write", _.ARGS.Val("<input type=\"hidden\" "));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("name=\"uxml_", aistaynum, "_", aithisreqmnt, "_", _.CALL(this, objunit, "UnitKey"), "\" ")));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("value=\"", _.CALL(this, _outer.server, "HtmlEncode", _.ARGS.Val(_.CALL(this, objunit, "EviivoMetaData"))), "\" />")));
            }

            if (_.IF(brendermaximumunitsavailable))
            {
                _.CALL(this, po, "Write", _.ARGS.Val("<div class=\"maxAvailUnits\">"));
                _.CALL(this, po, "Write", _.ARGS.Val("<p>"));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<span class=\"maxAvailUnitsLabelPrefix\">", _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/maxiumunitsavailableprefix").Val("Only ")), "</span>")));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<span class=\"maxAvailUnitsValue\">", _.CALL(this, objunit, "MaximumQuantityAvailable"), "</span>")));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<span class=\"maxAvailUnitsLabelSuffix\">", _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/maxiumunitsavailablesuffix").Val(" Rooms Remaining")), "</span>")));
                _.CALL(this, po, "Write", _.ARGS.Val("</p>"));
                _.CALL(this, po, "Write", _.ARGS.Val("</div>"));
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
        public object bookingui_renderbuttons(object aistaynum, object po, object bexternal)
        {
            object BookingUI_RenderButtons_retVal = null;
            object strclass = null;
            strclass = "btnBookStay";

            if (_.IF(bexternal))
            {
                strclass = _.CONCAT(strclass, " redirect");
            }

            _.CALL(this, po, "Write", _.ARGS.Val("<div class=\"pnStayButtons\">"));
            _.CALL(this, po, "Write", _.ARGS.Val("<input "));
            _.CALL(this, po, "Write", _.ARGS.Val("type=\"image\" "));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("class=\"", strclass, "\" ")));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("name=\"bookstay_", aistaynum, "\" ")));

            // Not using ids with VB Polling layout
            if (_.IF(_.NOT(_outer.isvbpollingenabled)))
            {
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("id=\"bookstay_", aistaynum, "\" ")));
            }

            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("value=\"", _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/btn/book").Val("Book")), "\" ")));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("src=\"", _.CALL(this, _outer.page, "ImageResource", _.ARGS.Val("bookonline/btn/book").Val(_.CONCAT(_.CALL(this, _outer.context, "ImageDir"), "booking/book.gif"))), "\" ")));
            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("alt=\"", _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/btn/book").Val("Book")), "\" />")));

            _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("</div>", VBScriptConstants.vbCrLf)));

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
        public object bookingui_availclassname(object abindicative, object abindicvalid, object abtelebook)
        {
            object BookingUI_AvailClassName_retVal = null;
            // If telephone booking, there's only one option
            if (_.IF(abtelebook))
            {
                BookingUI_AvailClassName_retVal = _.VAL(_.CALL(this, _outer, "BookingUI_AvailClassIcon", _.ARGS.Val("TELE")));
                return BookingUI_AvailClassName_retVal;
            }

            // If not telephone and not indicative, must be allocated
            if (_.IF(_.NOT(abindicative)))
            {
                BookingUI_AvailClassName_retVal = _.VAL(_.CALL(this, _outer, "BookingUI_AvailClassIcon", _.ARGS.Val("ALLOC")));
                return BookingUI_AvailClassName_retVal;
            }

            // Otherwise, get appropriate indicative option
            if (_.IF(abindicvalid))
            {
                BookingUI_AvailClassName_retVal = _.VAL(_.CALL(this, _outer, "BookingUI_AvailClassIcon", _.ARGS.Val("INDIC")));
            }
            else
            {
                BookingUI_AvailClassName_retVal = _.VAL(_.CALL(this, _outer, "BookingUI_AvailClassIcon", _.ARGS.Val("TELE")));
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
        public object bookingui_availclassicon(ref object asavailclassid)
        {
            object BookingUI_AvailClassIcon_retVal = null;
            object sicon = null;
            object stxt = null;
            object simg = null;

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

            simg = _.VAL(_.CALL(this, _outer.page, "ImageResource", _.ARGS.Val(_.CONCAT("bookonline/icons/", sicon)).Val(_.CONCAT(_.CALL(this, _outer.context, "ImageDir"), "booking/", sicon, ".gif"))));
            BookingUI_AvailClassIcon_retVal = _.CONCAT("<img src=\"", simg, "\" style=\"vertical-align:middle;\" alt=\"", stxt, "\" />");
            return BookingUI_AvailClassIcon_retVal;
        }

        // ====================================================================================================
        // RENDER: Format currency value
        // ====================================================================================================
        // SUMMARY: save space - only display price with pennies digits when fractional pounds
        // [amPrice]: money price to render
        // <retval>: string price
        public object bookingui_niceprice(ref object amprice)
        {
            object BookingUI_NicePrice_retVal = null;
            object strprice = null;
            // Get price:
            // - MakePrice will also handle any currency conversion)
            // - MakePrice will apply an appropriate currency symbol
            object byrefalias40 = amprice;
            try
            {
                strprice = _.VAL(_.CALL(this, _outer.page, "Functions", "Money", "MakePrice", _.ARGS.Ref(byrefalias40, v228 => { byrefalias40 = v228; })));
            }
            finally { amprice = byrefalias40; }

            // If there's a trailing ".00" then trim it off
            // NB: Pretty sure we'll never get a price of the form "?.00" - it should always
            //     be "?0.00", but just in case check that we've got a suitable long string
            if (_.IF(_.GT(_.NullableNUM(_.LEN(strprice)), (Int16)4)))
            {
                if (_.IF(_.EQ(_.NullableSTR(_.RIGHT(strprice, (Int16)3)), ".00")))
                {
                    strprice = _.VAL(_.LEFT(strprice, _.SUBT(_.LEN(strprice), (Int16)3)));
                }
            }

            // Return string ready for display
            BookingUI_NicePrice_retVal = _.VAL(_.CALL(this, _outer.server, "HTMLEncode", _.ARGS.Ref(strprice, v229 => { strprice = v229; })));
            return BookingUI_NicePrice_retVal;
        }

        // ====================================================================================================
        // RENDER: Pull description of linked unit (includes name of linked unit, name of source unit and
        // size of linked unit)
        // ====================================================================================================
        // SUMMARY: get description of linked unit - this is the PHYSICAL unit description
        public object bookingui_linkedunitdesc(object objunit, object objallunits)
        {
            object BookingUI_LinkedUnitDesc_retVal = null;
            object sunitname = null;
            object slinkedunitname = null;
            object objparentunit = null;
            object intindex = null;

            // If either UnitName of LinkedUnitName absent, return blank
            sunitname = _.VAL(_.CALL(this, objunit, "UnitName"));
            slinkedunitname = _.VAL(_.CALL(this, objunit, "LinkUnitName"));
            if (_.IF(_.OR(_.OR(_.OR(_.ISNULL(sunitname), _.EQ(_.NullableSTR(sunitname), "")), _.ISNULL(slinkedunitname)), _.EQ(_.NullableSTR(slinkedunitname), ""))))
            {
                BookingUI_LinkedUnitDesc_retVal = "";
                return BookingUI_LinkedUnitDesc_retVal;
            }

            // 2014-08-26 DWR: We need to retrieve the capacity of the unit that this linked unit is linked to. This data is not available in the avail
            // data from TOv2 since it is not included in the data from the Availability Component. It is why the "all units" data must be passed into
            // this method. This change addresses FogBugz 12998.
            objparentunit = VBScriptConstants.Nothing;
            var loopEnd28 = _.NUM(_.SUBT(_.CALL(this, objallunits, "Count"), (Int16)1));
            var loopStart28 = _.NUM((Int16)0, loopEnd28, (Int16)1);
            if (_.StrictLTE(loopStart28, loopEnd28))
            {
                for (intindex = loopStart28; _.StrictLTE(intindex, loopEnd28); intindex = _.ADD(intindex, (Int16)1))
                {
                    if (_.IF(_.EQ(_.CALL(this, _.CALL(this, objallunits, "getItem", _.ARGS.Ref(intindex, v230 => { intindex = v230; })), "Key"), _.CALL(this, objunit, "LinkUnitKey"))))
                    {
                        objparentunit = _.OBJ(_.CALL(this, objallunits, "getItem", _.ARGS.Ref(intindex, v231 => { intindex = v231; })));
                        break;
                    }
                }
            }
            if (_.IF(_.IS(objparentunit, VBScriptConstants.Nothing)))
            {
                _.CALL(this, _outer.page, "PrintTraceWarning", _.ARGS.Val(_.CONCAT("Unable to locate parent unit (", _.CALL(this, objunit, "LinkUnitKey"), ") for linked unit ", _.CALL(this, objunit, "UnitKey"))));
                BookingUI_LinkedUnitDesc_retVal = "";
                return BookingUI_LinkedUnitDesc_retVal;
            }

            BookingUI_LinkedUnitDesc_retVal = _.REPLACE(_.REPLACE(_.REPLACE(_.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/alsosoldaswithpersoncapacity").Val("(<i>#linkedunitname#</i> sold as #unitname# with #linkunitsize# person capacity)")), "#linkedunitname#", slinkedunitname), "#unitname#", sunitname), "#linkunitsize#", _.CALL(this, objparentunit, "Capacity"));
            return BookingUI_LinkedUnitDesc_retVal;
        }

        // ====================================================================================================
        // RENDER: This handles all of the rendering for ticketing - none of the StaySummary, StayDetails,
        // RenderButtons malarkey is required
        // ====================================================================================================
        public object bookingui_ticketssummary(ref object objavailentry, ref object adtstartnight, ref object po)
        {
            object BookingUI_TicketsSummary_retVal = null;
            object itotal = null;
            object isubtotal = null;
            object iselectedqty = null;
            object intindexunit = null;
            object objunit = null;
            object strpricebasis = null;

            if (_.IF(_.GT(_.NullableNUM(_.CALL(this, objavailentry, "Units", "Count")), (Int16)0)))
            {
                _.CALL(this, po, "Write", _.ARGS.Val("<div id=\"availabilityCalendarTableWrapper\">"));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<h3>", _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/ticketsavailable").Val("Tickets Available:")), "</h3>")));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<table id=\"availabilityCalendarTable\" summary=\"", _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/ticketsavailable").Val("Tickets Available")), "\" border=\"1\">")));
                _.CALL(this, po, "Write", _.ARGS.Val("<thead>"));
                _.CALL(this, po, "Write", _.ARGS.Val("<tr class=\"heading\">"));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<th class=\"unit\">", _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/tickets").Val("Tickets")), "</th>")));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<th class=\"select\">", _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/selection").Val("Selection")), "</th>")));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<th class=\"date\">", _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/date").Val("Date")), "</th>")));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<th class=\"total\">", _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/total").Val("Total")), "</th>")));
                _.CALL(this, po, "Write", _.ARGS.Val("</tr>"));
                _.CALL(this, po, "Write", _.ARGS.Val("<tr>"));
                _.CALL(this, po, "Write", _.ARGS.Val("<th></th>"));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<th class=\"number\">", _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/nooftickets").Val("No.Tickets")), "</th>")));
                object byrefalias41 = adtstartnight;
                try
                {
                    _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<th class=\"staydate\">", _.CALL(this, _outer.page, "Functions", "Dates", "NiceDateGuts", _.ARGS.Ref(byrefalias41, v232 => { byrefalias41 = v232; }).Val(true).Val(true)), "</th>")));
                }
                finally { adtstartnight = byrefalias41; }
                _.CALL(this, po, "Write", _.ARGS.Val("<th class=\"total\"></th>"));
                _.CALL(this, po, "Write", _.ARGS.Val("</tr>"));
                _.CALL(this, po, "Write", _.ARGS.Val("</thead>"));
                _.CALL(this, po, "Write", _.ARGS.Val("<tbody>"));
                itotal = (Int16)0;

                var loopEnd29 = _.NUM(_.SUBT(_.CALL(this, objavailentry, "Units", "Count"), (Int16)1));
                var loopStart29 = _.NUM((Int16)0, loopEnd29, (Int16)1);
                if (_.StrictLTE(loopStart29, loopEnd29))
                {
                    for (intindexunit = loopStart29; _.StrictLTE(intindexunit, loopEnd29); intindexunit = _.ADD(intindexunit, (Int16)1))
                    {
                        objunit = _.OBJ(_.CALL(this, objavailentry, "Units", "GetItem", _.ARGS.Ref(intindexunit, v234 => { intindexunit = v234; })));

                        iselectedqty = _.CLNG(_.CALL(this, _outer.request, "Form", _.ARGS.Val(_.CONCAT("unit_", _.CALL(this, objunit, "UnitKey")))));

                        if (_.IF(_.CALL(this, objunit, "PerPerson")))
                        {
                            strpricebasis = "per per";
                        }
                        else
                        {
                            strpricebasis = "per tic";
                        }

                        _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<tr id=\"row_", _.CALL(this, objunit, "UnitKey"), "\">")));
                        _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<td class=\"unit\">", _.CALL(this, objunit, "UnitName"), "</td>")));
                        _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<td class=\"select\">", _.CALL(this, _outer.page, "Functions", "Booking", "DrawSelectRange", _.ARGS.Val(_.CONCAT("unit_", _.CALL(this, objunit, "UnitKey"))).Val((Int16)0).Val(_.CALL(this, objunit, "UnitCount")).Ref(iselectedqty, v235 => { iselectedqty = v235; })), "</td>")));
                        _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<td class=\"price\">", _.CALL(this, _outer.server, "HTMLEncode", _.ARGS.Val(_.CALL(this, _outer.page, "Functions", "Money", "MakePrice", _.ARGS.Val(_.CALL(this, objunit, "StayTotalPayable"))))), "</td>")));
                        _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<td class=\"total\">", "<input type=\"hidden\" name=\"data_", _.CALL(this, objunit, "UnitKey"), "\" id=\"data_", _.CALL(this, objunit, "UnitKey"), "\" value=\"", _.CALL(this, objunit, "UnitCount"), ",", _.CALL(this, objunit, "MinOcc"), ",", _.CALL(this, objunit, "UnitSize"), ",", strpricebasis, ",", _.CALL(this, objunit, "StayTotalPayable"), "\">", _.CALL(this, _outer.server, "HTMLEncode", _.ARGS.Val(_.CALL(this, _outer.page, "Functions", "Money", "MakePrice", _.ARGS.Val(_.MULT(_.CALL(this, objunit, "StayTotalPayable"), iselectedqty))))), "</td>")));
                        _.CALL(this, po, "Write", _.ARGS.Val("</tr>"));
                        itotal = _.ADD(itotal, _.MULT(_.CALL(this, objunit, "StayTotalPayable"), iselectedqty));

                    }
                }
                isubtotal = _.ADD(isubtotal, itotal);

                _.CALL(this, po, "Write", _.ARGS.Val("</tbody>"));
                _.CALL(this, po, "Write", _.ARGS.Val("</table>"));
                _.CALL(this, po, "Write", _.ARGS.Val("</div>"));
                _.CALL(this, po, "Write", _.ARGS.Val("<table id=\"availabilityTotals\" summary=\"Totals\" border=\"1\">"));
                _.CALL(this, po, "Write", _.ARGS.Val("<tr>"));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<th>", _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/unitselection/grandtotal").Val("Grand Total")), "</th>")));
                _.CALL(this, po, "Write", _.ARGS.Val("<noscript>"));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<td><input type=\"image\" src=\"", _.CALL(this, _outer.page, "ImageResource", _.ARGS.Val("bookonline/unitselection/recalculate").Val(_.CONCAT(_.CALL(this, _outer.context, "ImageDir"), "booking/bookrecalculate.gif"))), "\" name=\"recalculate\" value=\"recalculate\" class=\"submit\"/></td>")));
                _.CALL(this, po, "Write", _.ARGS.Val("</noscript>"));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<td id=\"AvCalTotal\">", _.CALL(this, _outer.server, "HTMLEncode", _.ARGS.Val(_.CALL(this, _outer.page, "Functions", "Money", "MakePrice", _.ARGS.Ref(isubtotal, v237 => { isubtotal = v237; })))), "</td>")));
                _.CALL(this, po, "Write", _.ARGS.Val("</tr>"));
                _.CALL(this, po, "Write", _.ARGS.Val("</table>"));
                _.CALL(this, po, "Write", _.ARGS.Val(_.CONCAT("<input type=\"image\" src=\"", _.CALL(this, _outer.page, "ImageResource", _.ARGS.Val("bookonline/btn/bookticketing").Val(_.CONCAT(_.CALL(this, _outer.context, "ImageDir"), "booking/bookticketing.gif"))), "\" name=\"bookit\" value=\"", _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/btn/book").Val("Book")), "\" alt=\"", _.CALL(this, _outer.page, "Resource", _.ARGS.Val("bookonline/btn/book").Val("Book")), "\" class=\"submit\"/>")));
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
        public object getextbookurlfromproductestate(ref object asestateid)
        {
            object GetExtBookUrlFromProductEstate_retVal = null;
            object strposturl_ext = null;
            object strposturl_extdflt = null;
            object aryextbookestate = null;
            object i = null;
            // 2009-02-13 DWR: Can't remove spaces from content here because estate ids can have
            // spaces in (eg. "Arun DC" in TSE)
            aryextbookestate = _.SPLIT(_.REPLACE(_.TRIM(_.CONCAT("", _.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("Booking_ExtBookEstateMapping")))), VBScriptConstants.vbCrLf, ""), ",");
            var loopEnd30 = _.NUM(_.SUBT(_.UBOUND(aryextbookestate), (Int16)1));
            var loopStart30 = _.NUM((Int16)0, loopEnd30, (Int16)2);
            if (_.StrictLTE(loopStart30, loopEnd30))
            {
                for (i = loopStart30; _.StrictLTE(i, loopEnd30); i = _.ADD(i, (Int16)2))
                {
                    if (_.IF(_.EQ(_.NullableSTR(_.UCASE(_.TRIM(_.CALL(this, aryextbookestate, _.ARGS.Ref(i, v239 => { i = v239; }))))), "DEFAULT")))
                    {
                        strposturl_extdflt = _.VAL(_.CALL(this, aryextbookestate, _.ARGS.Val(_.ADD(i, (Int16)1))));
                    }
                    else if (_.IF(_.EQ(_.UCASE(_.TRIM(_.CALL(this, aryextbookestate, _.ARGS.Ref(i, v240 => { i = v240; })))), _.UCASE(_.TRIM(asestateid)))))
                    {
                        strposturl_ext = _.VAL(_.CALL(this, aryextbookestate, _.ARGS.Val(_.ADD(i, (Int16)1))));
                        break;
                    }
                }
            }

            if (_.IF(_.NOTEQ(_.NullableSTR(strposturl_ext), "")))
            {
                GetExtBookUrlFromProductEstate_retVal = _.VAL(strposturl_ext);
                _.CALL(this, _outer.page, "PrintTrace", _.ARGS.Val(_.CONCAT("GetExtBookUrlFromProductEstate: Product Estate ID = ", asestateid, ", External Book Url = ", strposturl_ext)));
            }
            else if (_.IF(_.NOTEQ(_.NullableSTR(strposturl_extdflt), "")))
            {
                GetExtBookUrlFromProductEstate_retVal = _.VAL(strposturl_extdflt);
                _.CALL(this, _outer.page, "PrintTrace", _.ARGS.Val(_.CONCAT("GetExtBookUrlFromProductEstate: Product Estate ID = ", asestateid, ", Using Default External Book Url = ", strposturl_extdflt)));
            }
            else
            {
                _.RAISEERROR(_.ADD(VBScriptConstants.vbObjectError, (Int16)1), "ETWP.Booking_UnitSelection Control", _.CONCAT("Failed to get External Booking Url [", asestateid, "]"));
            }

            return GetExtBookUrlFromProductEstate_retVal;
        }

        public object initexternalbookingsettings()
        {
            object InitExternalBookingSettings_retVal = null;
            if (_.IF(_.AND(_.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("Booking_ForceExternal")), _.NOTEQ(_.NullableSTR(_.TRIM(_.CONCAT("", _.CALL(this, _outer.page, "Site", "Params", _.ARGS.Val("Booking_ExtBookEstateMapping"))))), ""))))
            {
                _outer.isexternalbooking = true;
            }
            else
            {
                _outer.isexternalbooking = false;
            }
            return InitExternalBookingSettings_retVal;
        }

        // ====================================================================================================
        // MISC: Since the RenderSettings.BookingRequirement references passed into here are usually read-only
        // instances from the Page.Functions.GetSharedObject method, we'll need to make a local copy that we
        // can manipulate (since in some cases we need to mess about with the values)
        // ====================================================================================================
        public object geteditablebookingrequirement(object objbookingrequirement)
        {
            object GetEditableBookingRequirement_retVal = null;
            object objbookingrequirementnew = null;

            objbookingrequirementnew = _.OBJ(_.CALL(this, _outer.page, "Functions", "GetNewObject", _.ARGS.Val("BookingRequirement")));
            var with = _.OBJ(objbookingrequirementnew);
            _.SET(_.VAL(_.CALL(this, objbookingrequirement, "VisitDate")), this, with, "VisitDate");
            _.SET(_.VAL(_.CALL(this, objbookingrequirement, "Nights")), this, with, "Nights");
            _.SET(_.VAL(_.CALL(this, objbookingrequirement, "FlexibleRange")), this, with, "FlexibleRange");
            _.SET(_.VAL(_.CALL(this, objbookingrequirement, "Adults")), this, with, "Adults");
            _.SET(_.VAL(_.CALL(this, objbookingrequirement, "Children")), this, with, "Children");
            _.SET(_.VAL(_.CALL(this, objbookingrequirement, "ChildAges")), this, with, "ChildAges");
            _.SET(_.VAL(_.CALL(this, objbookingrequirement, "IsEviivoBooking")), this, with, "IsEviivoBooking");
            _.SET(_.VAL(_.CALL(this, objbookingrequirement, "Consumer")), this, with, "Consumer");
            _.SET(_.VAL(_.CALL(this, objbookingrequirement, "Offer")), this, with, "Offer");
            _.SET(_.VAL(_.CALL(this, objbookingrequirement, "BookingPassword")), this, with, "BookingPassword");
            _.SET(_.VAL(_.CALL(this, objbookingrequirement, "Product")), this, with, "Product");
            _.SET(_.VAL(_.CALL(this, objbookingrequirement, "Requirement")), this, with, "Requirement");
            _.SET(_.VAL(_.CALL(this, objbookingrequirement, "RequirementRef")), this, with, "RequirementRef");
            // NP 2012-03-12: RoomRequirements are needed
            // See GenerateRequirementFormData and Page.Functions.Booking.GenerateRequirementKeyValueData
            // the "NumRoomReq" value is part of the RoomRequirement, if it is not available then GenerateRequirementKeyValueData
            // sets default values for the adult and number of room requirements both to 1.
            // Requirements are not being passed to the RequirementSummary control correctly because the BookingRequestDictionary
            // is being overwritten with these incorrect default values.
            _.SET(_.VAL(_.CALL(this, objbookingrequirement, "RoomRequirements")), this, with, "RoomRequirements");
            GetEditableBookingRequirement_retVal = _.OBJ(objbookingrequirementnew);
            return GetEditableBookingRequirement_retVal;
        }
    }

    public sealed class EnvironmentReferences : EnvironmentReferencesBase
    {
        public object hlcontext { get => GetExternalReferenceAsObject(); internal set => RestoreExternalReferenceAsObject(value); }
    }
}