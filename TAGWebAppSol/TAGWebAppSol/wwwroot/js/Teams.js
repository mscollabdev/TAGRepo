//window.addEventListener(
//    "load",
//    () => {
//        init();      
//    },
//    false,
//);

var ddTeamLOption = [];
var allCustomApps = [];
var teamCustomApps = [];
var allGroupsReport = [];
var allUsers = [];
var UserID = "";
var userEmail = "";
var userDomain = "";
var siteName = "";
var siteID = "";
var webID = "";
var listID = "";



$(document).ready(function () {
    init();
    $("#ddlgroup").multiselect();

    $('#ulRequestTab li').click(function () {
        $('#reqtab-1').addClass('current');
    });

    $("#lblIsArchiveBtn").click(function () {
        $("#lblDetailsMessage").show();
        $("#lblDetailsMessage").text("Request raised succesfully! Please check Requests dashboard!");
    });

    $("#lblIsOrphanBtn").click(function () {
        $("#lblDetailsMessage").show();
        $("#lblDetailsMessage").text("Request raised succesfully! Please check Requests dashboard!");
    });


    setTimeout(function () {
        $("#ddlOwnerTeams").multiselect({
            search: true,
            selectAll: true,
            unselectAll: true
        });

        $("#ddlMemberTeams").multiselect({
            search: true,
            selectAll: true,
            unselectAll: true
        });

        $("[id*='opt']").click(function () {
            $("#lblMessage").hide();
            $("#lblMessageMember").hide();
            $("#lblDetailsMessage").hide();
            $("#txtUser").val('');
            if ($("[id*='opt']:checked").length === 1) {
                $("#divTeamTabDetails").show();
                $("#no-data").hide();
                $('#li-tab-2').removeClass('current');
                $("#tab-2").removeClass('current');
                $('#li-tab-1').addClass('current');
                $("#tab-1").addClass('current');
                $("#tab-2-inner").show();
                $("#tab-1-inner").hide();

                request
                    .getTeamsDetails($("[id*='opt']:checked")['0'].value)
                    .then((res3) => {
                        $("#tDetails").attr("href", res3.webUrl);
                        $("#tDetails").text(res3.displayName);
                        $("#lblIsArchive").text(res3.isArchived);
                        if (res3.isArchived === "true") {
                            $("#lblIsArchiveBtn").text("Restore Team");
                            $("#lblIsArchive").attr("style", "color:darkgreen;text-transform: uppercase;font-weight:bold");
                        }
                        else {
                            $("#lblIsArchiveBtn").text("Archive Team");
                            $("#lblIsArchive").attr("style", "color:darkred;text-transform: uppercase;font-weight:bold");
                        }
                        //$("#tName").text(res3.displayName);
                        //$("#tLink").html("<a href=" + res3.webUrl + ">Link</a>");

                    });

                request
                    .getTeamsChannel($("[id*='opt']:checked")['0'].value)
                    .then((res9) => {
                        $("#tChannels").text(res9.value.length);
                    });

                request
                    .getTeamsOwners($("[id*='opt']:checked")['0'].value, "test")
                    .then((res4) => {
                        if (res4.value.length) {
                            $("#lblIsOrphan").text("False").attr("style", "color:darkred;text-transform: uppercase;font-weight:bold");
                            $("#lblIsOrphan").append("<span tooltip='This team has " + res4.value.length + " Owner(s)'><b>*</b></span>");
                            $("#lblIsOrphanBtn").prop('disabled', true);
                        }
                        else {
                            $("#lblIsOrphan").text("True").attr("style", "color:darkgreen;text-transform: uppercase;font-weight:bold");
                            $("#lblIsOrphan").append("<span tooltip='This team has " + res4.value.length + " Owner(s). Please click Request for Owner button to raise a request for new team owner'><b>*</b></span>");
                            $("#lblIsOrphanBtn").prop('disabled', false);
                        }
                    });

                request
                    .getAppsDetails($("[id*='opt']:checked")['0'].value)
                    .then((res6) => {
                        for (var i = 0; i < res6.value.length; i++) {
                            var x = "'" + i + "'";
                            x = x.replace(/['"]+/g, '');
                            for (var j = 0; j < allCustomApps.length; j++) {
                                if (allCustomApps[j] === res6.value[x].teamsAppDefinition.teamsAppId) {
                                    teamCustomApps.push(res6.value[x].teamsAppDefinition.displayName);
                                }
                            }
                        }
                        $("#tApps").text(res6.value.length);
                        $("#tAppsOOTB").text(res6.value.length - teamCustomApps.length);
                        $("#tAppsCustom").text(teamCustomApps.length);
                    });
                       

                request
                    .getGroupNickName($("[id*='opt']:checked")['0'].value)
                    .then((res11) => {                       
                        siteName = res11.mailNickname;  
                        request
                            .getSiteWebID(userDomain, siteName)
                            .then((res12) => {
                                var siteDetails = res12.id;
                                siteID = siteDetails.split(',')[1];
                                webID = siteDetails.split(',')[2];

                                request
                                    .getSPList(siteID)
                                    .then((res13) => {
                                        for (var i = 0; i < res13.value.length; i++) {
                                            var x = "'" + i + "'";
                                            x = x.replace(/['"]+/g, '');
                                            if (res13.value[x].name === "Shared Documents") {
                                                listID = res13.value[x].id;
                                            }
                                        }
                                    });
                            });
                    });

               


               


                $.each(allGroupsReport, function (index, value) {
                    if (allGroupsReport[index].indexOf($("[id*='opt']:checked")['0'].value) > -1) {
                        $("#reportDate").text(value.split('|')[1]);
                        $("#tLastActivityDate").text(value.split('|')[2]);
                    }
                });

            }
            else if ($("[id*='opt']:checked").length === 0) {
                $("#divTeamTabDetails").hide();
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


        $("#sTOwner").text($("#ddlOwnerTeams option").length);
        $("#sTMember").text($("#ddlMemberTeams option").length);

        $(".ts-messages-header").css("display", "none");

    }, 2000);

    $('ul.tabs li').click(function () {
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
                $('#tab-1-inner').hide();
                $('#tab-2-inner').show();
                $("#no-data").hide();
            }
            else if ($("[id*='opt']:checked").length === 0) {
                $("#divTeamTabDetails").hide();
                $('#tab-1-inner').hide();
                $('#tab-2-inner').hide();
                $("#no-data").show();
            }
            else {
                $("#divTeamTabDetails").show();
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

    $(".form-check-input").click(function () {
        $("#lblMessage").hide();
        $("#lblMessageMember").hide();
        $("#lblDetailsMessage").hide();
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
            $("#divArchive").show();
            $("#divOrphan").hide();      
        }
        else if ($(this).attr("value") === "2") {
            $("#dTMembers").show();
            $("#dTOwners").hide();
            $("#dTUsers").hide();
            $("#divArchive").hide();
            $("#divOrphan").show();
        }
        else if ($(this).attr("value") === "3") {
            $("#no-data").hide();
            $("#dTMembers").hide();
            $("#dTOwners").hide();
            $("#divArchive").hide();
            $("#divOrphan").hide();
            $("#dTUsers").show();
        }
    });

    $("#btnLoad").click(function () {
        location.reload();
    });

    $("#addUserbtn").click(function () {
        $("#lblMessage").show();
        $("#lblMessageMember").hide();
        $("[id*='opt']:checked").each(function () {
            var groupId = $(this).val();
            request.
                addUser(groupId)
                .then((res10) => {
                    $("#lblMessage").text("User added successfully !");
                })
                .catch((error) => {
                    $("#lblMessage").text(error);
                });
        });
    });

    $("#leaveTeamsbtn").click(function () {
        $("#lblMessage").hide();
        $("#lblMessageMember").show();
        $("#txtUser").val('');
        $("[id*='opt']:checked").each(function () {
            var groupId = $(this).val();
            request.
                removeUser(UserID, groupId)
                .then((res18) => {
                    $("#lblMessageMember").text("User removed successfully! Please click 'Load Teams' again.");
                })
                .catch((error) => {
                    $("#lblMessageMember").text(error);
                    console.log(error);
                });
        });
    });

    $(".syncFiles").click(function () {
        window.open(
            "odopen://sync/?siteId=" + siteID + "&amp;webId=" + webID + "&amp;listId=" + listID + "&amp;userEmail=" + userEmail + "&amp;webUrl=https://" + userDomain + "/sites/" + siteName,
            "_blank" 
        );
    });

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





let client;
let client1;
let client2;
let client3;
let client4;
let client5;
let client6;
let client7;
let client8;
let client9;
let client10;
let client11;
let client12;
let client13;


let request = {
    getUserDetails: async () => {
        try {
            let res = await client.api("https://graph.microsoft.com/beta/me").get();
            return res;
        } catch (error) {
            throw error;
        }

    },

    getJoinedTeams: async () => {
        try {
            let res = await client1.api("https://graph.microsoft.com/beta/me/joinedTeams").get();
            return res;
        } catch (error) {
            throw error;
        }

    },

    getTeamsOwners: async (tID, tName) => {
        try {
            let res = await client2.api("https://graph.microsoft.com/beta/groups/" + tID + "/owners").get();
            res.tid = tID;
            res.tName = tName;
            return res;
        } catch (error) {
            throw error;
        }

    },

    getTeamsDetails: async (tID) => {
        try {
            let res = await client3.api("https://graph.microsoft.com/beta/teams/" + tID).get();
            return res;
        } catch (error) {
            throw error;
        }

    },

    getAppsDetails: async (tID) => {
        try {
            let res = await client4.api("https://graph.microsoft.com/beta/teams/" + tID + "/installedApps?$expand=teamsAppDefinition").get();
            return res;
        } catch (error) {
            throw error;
        }
    },

    getMemberDetails: async (tID) => {
        try {
            let res = await client5.api("https://graph.microsoft.com/beta/teams/" + tID + "/members").get();
            return res;
        } catch (error) {
            throw error;
        }
    },

    getTeamApps: async (tID) => {
        try {
            let res = await client6.api("https://graph.microsoft.com/beta/appCatalogs/teamsApps?$filter=distributionMethod eq 'organization'").get();
            return res;
        } catch (error) {
            throw error;
        }
    },

    getTeamsReport: async () => {
        try {
            let res = await client7.api("https://graph.microsoft.com/beta/reports/getOffice365GroupsActivityDetail(period='D90')?$format=application/json").get();
            return res;
        } catch (error) {
            throw error;
        }
    },

    getTeamsChannel: async (tID) => {
        try {
            let res = await client8.api("https://graph.microsoft.com/beta/teams/" + tID + "/channels").get();
            return res;
        } catch (error) {
            throw error;
        }
    },



    addUser: async (tID) => {
        try {
            var directoryObject = {};
            directoryObject["@odata.id"] = "https://graph.microsoft.com/beta/directoryObjects/9fd6a31b-5490-42cb-a669-57b1530b5771";
            let res = await client9.api("https://graph.microsoft.com/beta/groups/" + tID + "/members/$ref").post(directoryObject);
            return res;
        }
        catch (error) {
            throw error;
        }
    },

    removeUser: async (uID, tID) => {
        try {
            let res = await client10.api("https://graph.microsoft.com/beta/groups/" + tID + "/members/" + uID + "/$ref").delete();
            return res;
        } catch (error) {
            throw error;
        }
    },

    getGroupNickName: async (tID) => {
        try {
            let resp = await client11.api("https://graph.microsoft.com/beta/groups/" + tID + "?$select=mailNickname").get();
            return resp;
        } catch(error) {
            throw error;
        }
    },

    getSiteWebID: async (userDomain, siteName) => {
        try {
            let res = await client12.api("https://graph.microsoft.com/beta/sites/" + userDomain + ":/sites/" + siteName + "?$select=id").get();
            return res;
        } catch (error) {
            throw error;
        }
    },

    getSPList: async (siteid) => {
        try {
            let res = await client13.api("https://graph.microsoft.com/beta/sites/" + siteid + "/lists?$select=id,name").get();
            return res;
        } catch (error) {
            throw error;
        }
    }


};


const init = async () => {

    const scopes = ["Group.ReadWrite.All"];
    const CacheLocation = "localStorage";
    const msalConfig = {
        auth: {
            clientId: "2f625ade-9dc8-414f-a0ea-4efe97780e30",
            authority: "https://login.microsoftonline.com/common"
        },
        cache: {
            cacheLocation: CacheLocation
        }
    };


    var email = "";
    var teamIDs = [];
    var msalApplication = new Msal.UserAgentApplication(msalConfig);
    const msalOptions = new MicrosoftGraph.MSALAuthenticationProviderOptions(scopes);
    const msalProvider = new MicrosoftGraph.ImplicitMSALAuthenticationProvider(msalApplication, msalOptions);



    client = MicrosoftGraph.Client.initWithMiddleware({
        debugLogging: true,
        authProvider: msalProvider
    });    
    client1 = MicrosoftGraph.Client.initWithMiddleware({
        debugLogging: true,
        authProvider: msalProvider
    });
    client2 = MicrosoftGraph.Client.initWithMiddleware({
        debugLogging: true,
        authProvider: msalProvider
    });
    client3 = MicrosoftGraph.Client.initWithMiddleware({
        debugLogging: true,
        authProvider: msalProvider
    });
    client4 = MicrosoftGraph.Client.initWithMiddleware({
        debugLogging: true,
        authProvider: msalProvider
    });
    client5 = MicrosoftGraph.Client.initWithMiddleware({
        debugLogging: true,
        authProvider: msalProvider
    });
    client6 = MicrosoftGraph.Client.initWithMiddleware({
        debugLogging: true,
        authProvider: msalProvider
    });
    client7 = MicrosoftGraph.Client.initWithMiddleware({
        debugLogging: true,
        authProvider: msalProvider
    });
    client8 = MicrosoftGraph.Client.initWithMiddleware({
        debugLogging: true,
        authProvider: msalProvider
    });
    client9 = MicrosoftGraph.Client.initWithMiddleware({
        debugLogging: true,
        authProvider: msalProvider
    });
    client10 = MicrosoftGraph.Client.initWithMiddleware({
        debugLogging: true,
        authProvider: msalProvider
    });
    client11 = MicrosoftGraph.Client.initWithMiddleware({
        debugLogging: true,
        authProvider: msalProvider
    });
    client12 = MicrosoftGraph.Client.initWithMiddleware({
        debugLogging: true,
        authProvider: msalProvider
    });
    client13 = MicrosoftGraph.Client.initWithMiddleware({
        debugLogging: true,
        authProvider: msalProvider
    });
    client14 = MicrosoftGraph.Client.initWithMiddleware({
        debugLogging: true,
        authProvider: msalProvider
    });

    request
        .getUserDetails()
        .then((res) => {

            $("#sName").text(res.displayName);
            $("#sEmail").text(res.mail);
            $("#sJob").text(res.jobTitle);
            $("#sContact").text(res.mobilePhone);
            email = res.mail;
            UserID = res.id;
            userEmail = res.mail;
            userDomain = userEmail.split('@')[1];
            userDomain = userDomain.replace("onmicrosoft", "sharepoint");
        })
        .catch((error) => {
            console.log(error);
        });

    request
        .getTeamApps()
        .then((res5) => {
            for (var k = 0; k < res5.value.length; k++) {
                var z = "'" + k + "'";
                z = z.replace(/['"]+/g, '');
                allCustomApps.push(res5.value[z].id);
            }
        });

    request
        .getTeamsReport()
        .then((res7) => {
            for (var n = 0; n < res7.value.length; n++) {
                var t = "'" + n + "'";
                t = t.replace(/['"]+/g, '');
                allGroupsReport.push(res7.value[t].groupId + "|" + res7.value[t].reportRefreshDate + "|" + res7.value[t].lastActivityDate);
            }
        });


    request
        .getJoinedTeams()
        .then((res1) => {
            for (var i = 0; i < res1.value.length; i++) {
                var x = "'" + i + "'";
                x = x.replace(/['"]+/g, '');
                teamIDs.push(res1.value[x].id + "|" + res1.value[x].displayName);
            }

            meow();
        })
        .catch((error) => {
            console.log(error);
        });

    function meow() {
        var sOwners = $("<select id='ddlOwnerTeams' multiple name='dTO' />");
        var sMembers = $("<select id='ddlMemberTeams' multiple name='dTM' />");
        for (var a = 0; a < teamIDs.length; a++) {
            $("#sTeams").text(teamIDs.length);
            var teamID = [];
            teamID = teamIDs[a].split('|');
            request
                .getTeamsOwners(teamID[0], teamID[1])
                .then((res2) => {

                    var isOwner = false;

                    res2.value.forEach(function (val, i) {

                        if (email.indexOf(res2.value[i].mail) > -1) {
                            isOwner = true;
                        }
                    });

                    if (isOwner === true) {
                        $('<option />', { value: res2.tid, text: res2.tName }).appendTo(sOwners);
                    }
                    else {
                        $('<option />', { value: res2.tid, text: res2.tName }).appendTo(sMembers);
                    }

                });
        }

        sOwners.appendTo($("#dTOwners"));
        sMembers.appendTo($("#dTMembers"));

    }



};


