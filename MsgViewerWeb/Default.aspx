<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeBehind="Default.aspx.cs" Inherits="MsgViewerWeb._Default" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
    
        <script type="text/javascript">

            $(function () {

                $('#messages').find('a').each(function () {

                    $('#messageList').append('<li><a href="' + this.href + '">' + $(this).text() +'</a></li>');

                });

                $('#messageList').find('a').click(function (event) {
                    event.preventDefault();
                    window.open($(this).attr('href'));
                });

                $('#messages').hide();
            });

    </script>

</asp:Content>

<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    
    <ul id="messageList">
        
    </ul>
    <asp:Panel ID="messages" runat="server" ClientIDMode="Static" />

</asp:Content>
