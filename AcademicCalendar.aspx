<%@ Page Title="Academic Calendar" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="AcademicCalendar.aspx.vb" Inherits="AcademicCalendarProject.AcademicCalendar" %>

<asp:Content ID="ContentHead" ContentPlaceHolderID="HeadContent" runat="server">
    <link href="<%= ResolveUrl("~/CSS/AcademicCalendar.css") %>" rel="stylesheet" />
</asp:Content>
<asp:Content ID="Content1" ContentPlaceHolderID="HeaderContent" runat="server">
	 <div class="logo-banner">
			<div class="page-title">
				<h1>Academic Calendar</h1>
				<p>View academic events, deadlines, exams, registration dates, and holidays</p>
			</div>
		 <!--<img src="<%= ResolveUrl("~/Images/Sagesse.png") %>" class="page-logo" alt="Université La Sagesse Logo" />-->
		 <img src="<%= ResolveUrl("~/Images/ULS-logo-vertical.png") %>" class="page-logo xs-none" alt="Université La Sagesse Logo" />
		 <img src="<%= ResolveUrl("~/Images/ULS-logo-mobile.png") %>" class="page-logo xs-block" alt="Université La Sagesse Logo" />
	 </div>
</asp:Content>
<asp:Content ID="ContentMain" ContentPlaceHolderID="MainContent" runat="server">
	<asp:HiddenField ID="hfViewportWidth" runat="server" />
    <div class="calendar-wrapper">
		<div class="cal-controls">
			<div class="view-switch">
				<asp:Button ID="btnListView" runat="server" Text="List" CssClass="view-btn list-view-btn view-btn-active" />
				<asp:Button ID="btnCalendarView" runat="server" Text="Calendar" CssClass="view-btn calendar-view-btn" />
			</div>
			<div class="outlook-download-box">
				<!--<asp:Button ID="btnDownloadICS" runat="server" Text="📅 Download Calendar File" CssClass="outlook-download-btn" />-->
				<asp:HyperLink ID="lnkSubscribeOutlook" runat="server" Text="Subscribe in Outlook Calendar" CssClass="outlook-subscribe-link">
					Subscribe in Outlook Calendar
					<span class="btn-icon" aria-hidden="true">→</span>
				</asp:HyperLink>
			</div>
		</div>
        

        <div class="filter-box">
            <asp:LinkButton ID="lnkAll" runat="server" CssClass="filter-link"> <span class="dot dot-all"></span>All </asp:LinkButton>
            <asp:LinkButton ID="lnkExams" runat="server" CssClass="filter-link"> <span class="dot dot-exams"></span>Exams </asp:LinkButton>
            <asp:LinkButton ID="lnkDeadlines" runat="server" CssClass="filter-link"> <span class="dot dot-deadlines"></span>Deadlines </asp:LinkButton>
            <asp:LinkButton ID="lnkRegistration" runat="server" CssClass="filter-link"> <span class="dot dot-registration"></span>Registration</asp:LinkButton>
            <asp:LinkButton ID="lnkHolidays" runat="server" CssClass="filter-link"> <span class="dot dot-holidays"></span>Holidays</asp:LinkButton>
			<asp:LinkButton ID="lnkAcademic" runat="server" CssClass="filter-link"> <span class="dot dot-academic"></span>Academic</asp:LinkButton>
        </div>

        <asp:Label ID="lblError" runat="server" CssClass="error-message"></asp:Label>

        <asp:Panel ID="pnlListView" runat="server" CssClass="list-view-panel is-active-view">

            <asp:Literal ID="litListEvents" runat="server"></asp:Literal>

            <asp:Label ID="lblListMessage" runat="server" CssClass="message"></asp:Label>

        </asp:Panel>

        <asp:Panel ID="pnlCalendarView" runat="server" CssClass="calendar-view-panel">

            <div class="calendar-card">

                <div class="calendar-nav">
                    <asp:Button ID="btnPrevMonth" runat="server" Text="‹" CssClass="nav-btn" />

                    <div class="calendar-date">
                        <span class="calendar-month-title">
                            <asp:Literal ID="litCalendarMonth" runat="server"></asp:Literal>
                        </span>

                        <span class="calendar-year">
                            <asp:Literal ID="litCalendarYear" runat="server"></asp:Literal>
                        </span>
                    </div>

                    <asp:Button ID="btnNextMonth" runat="server" Text="›" CssClass="nav-btn" />
                </div>

                <asp:Literal ID="litCalendar" runat="server"></asp:Literal>

            </div>

        </asp:Panel>
		
    </div>

    <%--script that collapse all month cards after pressing list button--%>

    <%--<script type="text/javascript">
        function toggleMonthCard(bodyId, iconId) {
            var body = document.getElementById(bodyId);
            var icon = document.getElementById(iconId);

            if (body.style.display === "none") {
                body.style.display = "block";
                icon.innerHTML = "−";
            } else {
                body.style.display = "none";
                icon.innerHTML = "+";
            }
        }
	</script>--%>

<%--keeping current month expanded by default and other months collapsed by default--%>
<script type="text/javascript">
	function toggleMonthCard(bodyId, iconId) {
		var body = document.getElementById(bodyId);
		var icon = document.getElementById(iconId);

		if (!body || !icon) {
			return;
		}

		var storageKey = "monthCardState_" + bodyId;

        if (body.style.display === "none") {
            body.style.display = "flex";
            icon.innerHTML = "−";
            localStorage.setItem(storageKey, "expanded");
        } else {
            body.style.display = "none";
          
			icon.innerHTML = "+";
			localStorage.setItem(storageKey, "collapsed");
		}
	}

	function restoreMonthCardStates() {
		var bodies = document.getElementsByClassName("monthly-list-body");

		for (var i = 0; i < bodies.length; i++) {
			var body = bodies[i];
			var bodyId = body.id;

			if (!bodyId) {
				continue;
			}

			var iconId = bodyId.replace("monthBody_", "monthIcon_");
			var icon = document.getElementById(iconId);

			if (!icon) {
				continue;
			}

			var storageKey = "monthCardState_" + bodyId;
			var savedState = localStorage.getItem(storageKey);
			var defaultState = body.getAttribute("data-default-state");
			var isCurrentMonth = body.getAttribute("data-current-month");

			/*
			   Current month should be expanded by default.
			   We ignore old saved collapsed state for the current month.
			*/
			if (isCurrentMonth === "yes") {
				body.style.display = "flex";
				icon.innerHTML = "−";
				localStorage.setItem(storageKey, "expanded");
				continue;
			}

			if (savedState === "collapsed") {
				body.style.display = "none";
				icon.innerHTML = "+";
			} else if (savedState === "expanded") {
				body.style.display = "flex";
				icon.innerHTML = "−";
			} else {
				if (defaultState === "collapsed") {
					body.style.display = "none";
					icon.innerHTML = "+";
				} else {
					body.style.display = "flex";
					icon.innerHTML = "−";
				}
			}
		}
	}

	window.onload = function () {
    restoreMonthCardStates();
};
(function() {
    var field = document.getElementById("<%= hfViewportWidth.ClientID %>");
    if (!field) {
        return;
    }

    function currentWidth() {
        return window.innerWidth || document.documentElement.clientWidth || 0;
    }

    function syncViewportWidth() {
        field.value = String(currentWidth());
    }

    syncViewportWidth();
    window.addEventListener("resize", syncViewportWidth);
    document.addEventListener("submit", syncViewportWidth, true);

    var isMobile = currentWidth() <= 768;
    var serverIsMobile = <%= If(IsMobileRequest(), "true", "false") %>;

    if (isMobile !== serverIsMobile && typeof __doPostBack === "function") {
        syncViewportWidth();
        __doPostBack("<%= hfViewportWidth.UniqueID %>", "");
    }
})();


</script>

</asp:Content>