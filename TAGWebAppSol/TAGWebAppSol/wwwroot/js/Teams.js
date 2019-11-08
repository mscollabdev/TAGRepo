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



$(document).ready(function () {
    init();
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
            if ($("[id*='opt']:checked").length === 1) {
                var cID = $(this).id;               
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
                        $("#tName").text(res3.displayName);
                        $("#tLink").html("<a href=" + res3.webUrl + ">Link</a>");

                    });

                request
                    .getTeamsChannel($("[id*='opt']:checked")['0'].value)
                    .then((res9) => {
                        $("#tChannels").text(res9.value.length);
                    });

                request
                    .getTeamsOwners($("[id*='opt']:checked")['0'].value, "test")
                    .then((res4) => {
                        $("#tOwners").text(res4.value.length);
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
                        $("#tAppsOOTB").text(res6.value.length - teamCustomApps.length);
                        $("#tAppsCustom").text(teamCustomApps.length);
                    }); 

                $.each(allGroupsReport, function (index, value)   
	                {                    	                   
                    if (allGroupsReport[index].indexOf($("[id*='opt']:checked")['0'].value) > -1) {                                               
                        $("#reportDate").text(value.split('|')[1]);
                        $("#tLastActivityDate").text(value.split('|')[2]);
                    }
	                });

            }
            else if ($("[id*='opt']:checked").length === 0)
            {
                $("#divTeamTabDetails").hide();
                $("#no-data").show();
            }
            else {
                $("#no-data").hide();
                $('#li-tab-2').addClass('current');
                $("#tab-2").addClass('current');
                $('#li-tab-1').removeClass('current');
                $("#tab-1").removeClass('current');                               
            }
        });


        $("#sTOwner").text($("#ddlOwnerTeams option").length);
        $("#sTMember").text($("#ddlMemberTeams option").length);

    }, 1000);

    $('ul.tabs li').click(function () {
        var tab_id = $(this).attr('data-tab');
        console.log(tab_id);
        $('ul.tabs li').removeClass('current');
        $('.tab-content').removeClass('current');

        $(this).addClass('current');
        $("#" + tab_id).addClass('current');

        if (tab_id === "tab-1") {
            if ($(".form-check-input").attr("value") === "1") {
                $("#tab-1-inner").hide();
                $("#tab-2-inner").show();
            }
            else {
                $("#tab-1-inner").show();
                $("#tab-2-inner").hide();
            }
        }
    });

    $(".form-check-input").click(function () {
        if ($(this).attr("value") === "1") {
            $("#dTOwners").show();
            $("#dTMembers").hide();
        }
        if ($(this).attr("value") === "2") {
            $("#dTMembers").show();
            $("#dTOwners").hide();
        }
    });  

    $("#btnLoad").click(function () {
        location.reload();
    });

    $("#addUserbtn").click(function () {
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
    catch(error) {
        throw error;
    }
}

    
};


const init = async () => {
    
    const scopes = ["Group.ReadWrite.All"];
    const msalConfig = {
        auth: {
            clientId: "9a0a1545-40f3-4068-8058-2bf10b52fa18",
            authority: "https://login.microsoftonline.com/common"
        },
        cache: {
            cacheLocation: "localStorage",
            storeAuthStateInCookie: true
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

    request
        .getUserDetails()
        .then((res) => {
            $("#sName").text(res.displayName);
            $("#sEmail").text(res.mail);
            $("#sJob").text(res.jobTitle);
            $("#sContact").text(res.mobilePhone);
            email = res.mail;
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
                var x = "'"+ i +"'";
                x = x.replace(/['"]+/g, '');
                teamIDs.push(res1.value[x].id + "|" +res1.value[x].displayName);                 
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


