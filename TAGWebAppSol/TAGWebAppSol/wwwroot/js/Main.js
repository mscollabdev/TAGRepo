$(document).ready(function () {

    $("#tblUserDetails").kendoGrid({
        dataSource: {
            data: userTeams,
            schema: {
                model: {
                    fields: {
                        TeamName: { type: "string" },
                        UserType: { type: "string" }
                    }
                }
            },
            pageSize: 10
        },
        height: 300,
        scrollable: true,
        sortable: true,        
        filterable: true,
        pageable: {
            input: true,
            numeric: false
        },
        columns: [
            { field: "TeamName", title: "Team Name", width: "250px" },
            { field: "UserType", title: "User Type", width: "150px" },
            { command: ["destroy"], title: "&nbsp;", width: "100px" }
        ],
        editable: "inline"
    });

    $("#tblRequest").kendoGrid({
        dataSource: {
            data: userRequests,
            schema: {
                model: {
                    fields: {
                        TeamName: { type: "string" },
                        RequestType: { type: "string" },
                        DateRaised: { type: "string" },
                        Status: { type: "string" }
                    }
                }
            },
            pageSize: 10
        },
        height: 300,
        scrollable: true,
        sortable: true,
        filterable: true,
        pageable: {
            input: true,
            numeric: false
        },
        columns: [
            { field: "TeamName", title: "Team Name", width: "120px" },
            { field: "RequestType", title: "Request Type", width: "100px" },
            { field: "DateRaised", title: "Request Date", width: "100px" },
            { field: "Status", title: "Status", width: "100px" },
            {
                command: [{ name: "Details" }], title: "&nbsp;", width: "80px"
            }],            
        editable: "inline"
    });

    $("#ddlgroup").multiselect();

    $("#ddlOwnerTeams").multiselect({
        search: true,
        selectAll: true        
    });

    $("#ddlMemberTeams").multiselect({
        search: true,
        selectAll: true
    });

    $("#addUserbtn").click(function () {
        $("#lblMessage").text("User added successfully !");
    });

    $("#leaveTeamsbtn").click(function () {
        $("#lblMessageMember").text("User removed successfully! Please click 'Load Teams' again.");
    });

    $(".form-check-input").click(function () {
        $("#lblMessage").hide();
        $("#lblMessageMember").hide();
        $("#txtUser").val('');
        $("[id *= 'opt']").prop("checked", false);
        /* $('.ms-options-wrap').find('> button:first-child').text('');*/

        $("#divTeamTabDetails").hide();
        $("#divUserDetails").hide();
        $("#no-data").show();


        if ($(this).attr("value") === "1") {
            $("#dTOwners").show();
            $("#dTMembers").hide();
            $("#dTUsers").hide();
        }
        else if ($(this).attr("value") === "2") {
            $("#dTMembers").show();
            $("#dTOwners").hide();
            $("#dTUsers").hide();
        }
        else if ($(this).attr("value") === "3") {
            $("#no-data").hide();
            $("#dTMembers").hide();
            $("#dTOwners").hide();
            $("#dTUsers").show();
        }
    });

    $('#ulRequestTab li').click(function () {
        $('#reqtab-1').addClass('current');
    });

    $('#ulTeamTabDetails li').click(function () {
        var tab_id = $(this).attr('data-tab');
        $('ul.tabs li').removeClass('current');
        $('.tab-content').removeClass('current');
        $("#lblMessage").hide();
        $("#lblMessageMember").hide();
        $("#txtUser").val('');

        $(this).addClass('current');
        $("#" + tab_id).addClass('current');

        if (tab_id === "tab-1") {
            if ($("[id*='opt']:checked").length === 1) {
                $("#divTeamTabDetails").show();
                $("#divUserDetails").hide();
                $('#tab-1-inner').hide();
                $('#tab-2-inner').show();
                $("#no-data").hide();
            }
            else if ($("[id*='opt']:checked").length === 0) {
                $("#divTeamTabDetails").hide();
                $("#divUserDetails").hide();
                $('#tab-1-inner').hide();
                $('#tab-2-inner').hide();
                $("#no-data").show();
            }
            else {
                $("#divTeamTabDetails").show();
                $("#divUserDetails").hide();
                $('#tab-1-inner').show();
                $('#tab-2-inner').hide();
                $("#no-data").hide();
            }
        }
        else if (tab_id === "tab-2") {
            if ($("input[name='inlineRadioOptions']:checked").val() === "1") {
                $('#button-owner').show();
                $('#button-member').hide();
            }
            else if ($("input[name='inlineRadioOptions']:checked").val() === "2") {
                $('#button-owner').hide();
                $('#button-member').show();
            }

        }
               
       
    });

    $("[id*='opt']").click(function () {
        $("#divUserDetails").hide();
        $("#lblMessage").hide();
        $("#lblMessageMember").hide();
        if ($("[id*='opt']:checked").length === 1) {
            $("#divTeamTabDetails").show();
            $("#no-data").hide();
            $("#divUserDetails").hide();
            $('#li-tab-2').removeClass('current');
            $("#tab-2").removeClass('current');
            $('#li-tab-1').addClass('current');
            $("#tab-1").addClass('current');
            $("#tab-2-inner").show();
            $("#tab-1-inner").hide();
        }
        else if ($("[id*='opt']:checked").length === 0) {
            $("#divTeamTabDetails").hide();
            $("#divUserDetails").hide();
            $("#no-data").show();
        }
        else {
            if ($("input[name='inlineRadioOptions']:checked").val() === "1") {
                $('#button-owner').show();
                $('#button-member').hide();
            }
            else if ($("input[name='inlineRadioOptions']:checked").val() === "2") {
                $('#button-owner').hide();
                $('#button-member').show();
            }
            $("#no-data").hide();
            $('#li-tab-2').addClass('current');
            $("#tab-2").addClass('current');
            $('#li-tab-1').removeClass('current');
            $("#tab-1").removeClass('current');
        }
    });

    $('#searchUserbtn').click(function () {
        $("#no-data").hide();
        $("#lblMessage").hide();
        $("#lblMessageMember").hide();
        $("#divTeamTabDetails").hide();
        $("#divUserDetails").show();
        $("#tab-1").removeClass('current');
        $("#tab-2").removeClass('current');
        $("#tab-4").addClass('current');
    });
   
});