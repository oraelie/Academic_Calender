<%@ Page Title="Academic Calendar" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="AcademicCalendar.aspx.vb" Inherits="AcademicCalendarProject.AcademicCalendar" %>

<asp:Content ID="ContentHead" ContentPlaceHolderID="HeadContent" runat="server">
    <link href="<%= ResolveUrl("~/CSS/AcademicCalendar.css") %>" rel="stylesheet" />
</asp:Content>

<asp:Content ID="ContentMain" ContentPlaceHolderID="MainContent" runat="server">

    <div class="calendar-wrapper">

        <div class="logo-banner">
            <img src="<%= ResolveUrl("~/Images/Sagesse.png") %>" class="page-logo" alt="Université La Sagesse Logo" />
        </div>

        <div class="page-title">
            <h1>Academic Calendar</h1>
            <p>View academic events, deadlines, exams, registration dates, and holidays</p>
        </div>

        <div class="view-switch">
            <asp:Button ID="btnListView" runat="server" Text="☰  List" CssClass="view-btn view-btn-active" />
            <asp:Button ID="btnCalendarView" runat="server" Text="🟩  Calendar" CssClass="view-btn" />
        </div>

        <div class="filter-box">
            <span class="filter-title">FILTER</span>

            <asp:LinkButton ID="lnkAll" runat="server" CssClass="filter-link">
                <span class="dot dot-all"></span>All
            </asp:LinkButton>

            <asp:LinkButton ID="lnkExams" runat="server" CssClass="filter-link">
                <span class="dot dot-exams"></span>Exams
            </asp:LinkButton>

            <asp:LinkButton ID="lnkDeadlines" runat="server" CssClass="filter-link">
                <span class="dot dot-deadlines"></span>Deadlines
            </asp:LinkButton>

            <asp:LinkButton ID="lnkRegistration" runat="server" CssClass="filter-link">
                <span class="dot dot-registration"></span>Registration
            </asp:LinkButton>

            <asp:LinkButton ID="lnkHolidays" runat="server" CssClass="filter-link">
                <span class="dot dot-holidays"></span>Holidays
            </asp:LinkButton>

            <asp:LinkButton ID="lnkAcademic" runat="server" CssClass="filter-link">
                <span class="dot dot-academic"></span>Academic
            </asp:LinkButton>
        </div>

        <asp:Label ID="lblError" runat="server" CssClass="error-message"></asp:Label>

        <asp:Panel ID="pnlListView" runat="server">

            <asp:Literal ID="litListEvents" runat="server"></asp:Literal>

            <asp:Label ID="lblListMessage" runat="server" CssClass="message"></asp:Label>

        </asp:Panel>

        <asp:Panel ID="pnlCalendarView" runat="server">

            <div class="calendar-card">

                <div class="calendar-nav">
                    <asp:Button ID="btnPrevMonth" runat="server" Text="‹" CssClass="nav-btn" />

                    <div>
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


    <%--keeping the state of collapse and expand as the same state when pressing list button--%>
    <script type="text/javascript">
		function toggleMonthCard(bodyId, iconId) {
			var body = document.getElementById(bodyId);
			var icon = document.getElementById(iconId);

			if (!body || !icon) {
				return;
			}

			var storageKey = "monthCardState_" + bodyId;

			if (body.style.display === "none") {
				body.style.display = "block";
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

				var iconId = bodyId.replace("monthBody", "monthIcon");
				var icon = document.getElementById(iconId);

				if (!icon) {
					continue;
				}

				var storageKey = "monthCardState_" + bodyId;
				var savedState = localStorage.getItem(storageKey);

				if (savedState === "collapsed") {
					body.style.display = "none";
					icon.innerHTML = "+";
				} else {
					body.style.display = "block";
					icon.innerHTML = "−";
				}
			}
		}

		window.onload = function () {
			restoreMonthCardStates();
		};
	</script>


</asp:Content>