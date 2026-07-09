<%@ Page Title="Academic Calendar" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="AcademicCalendar.aspx.vb" Inherits="AcademicCalendarProject.AcademicCalendar" %>

<asp:Content ID="ContentHead" ContentPlaceHolderID="HeadContent" runat="server">
    <link href="<%= ResolveUrl("~/CSS/AcademicCalendar.css") %>" rel="stylesheet" />
</asp:Content>

<asp:Content ID="ContentMain" ContentPlaceHolderID="MainContent" runat="server">

    <div class="calendar-wrapper">

        <div class="page-title">
            <h1>Academic Calendar</h1>
            <p>Student Edition · Data loaded from Excel file</p>
        </div>

        <div class="view-switch">
            <asp:Button ID="btnListView" runat="server" Text="☰  List" CssClass="view-btn" />
            <asp:Button ID="btnCalendarView" runat="server" Text="▣  Calendar" CssClass="view-btn" />
        </div>

        <div class="filter-box">
            <span class="filter-title">FILTER:</span>

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
            <div class="list-card">

                <div class="month-header">
                    <asp:Literal ID="litListMonthTitle" runat="server"></asp:Literal>
                </div>

                <asp:Repeater ID="rptListEvents" runat="server">
                    <ItemTemplate>
                        <div class="list-row">
                            <div class="event-date">
                                <%# Eval("DayText") %>
                            </div>

                            <div>
                                <span class='<%# Eval("DotClass") %>'></span>
                            </div>

                            <div class="event-title">
                                <%# Eval("EventTitle") %>
                            </div>
                        </div>
                    </ItemTemplate>
                </asp:Repeater>

                <asp:Label ID="lblListMessage" runat="server" CssClass="message"></asp:Label>

            </div>
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

</asp:Content>