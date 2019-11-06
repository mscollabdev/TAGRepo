//window.addEventListener(
//    "load",
//    () => {
//        init();      
//    },
//    false,
//);

var ddTeamLOption = [];

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
                console.log(cID);
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
                        $("#tDesc").text(res3.description);
                        $("#tLink").html("<a href=" + res3.webUrl + ">Link</a>");
                        $("#tArchive").text(res3.isArchived);                        
                    });

                request
                    .getTeamsOwners($("[id*='opt']:checked")['0'].value, "test")
                    .then((res4) => {
                        $("#tOwners").text(res4.value.length);
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

              
});





let client;
let client1;
let client2;
let client3;

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

    }


};


const init = async () => {   
    const scopes = ["user.read", "profile", "User.ReadWrite", "Files.Read", "Files.Read.All", "Files.ReadWrite", "Files.ReadWrite.All", "Mail.Read", "Mail.ReadWrite", "Mail.Send"];
    const msalConfig = {
        auth: {
            clientId: "9a0a1545-40f3-4068-8058-2bf10b52fa18",
            redirectUri: "https://localhost:44373/"
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
            var teamID = [];            
            teamID = teamIDs[a].split('|');
            request
                .getTeamsOwners(teamID[0], teamID[1])                
                .then((res2) => {       
                    var count = 0;
                    for (var j = 0; j < res2.value.length; j++) {                        
                        var y = "'" + j + "'";
                        y = y.replace(/['"]+/g, '');

                        if(email === res2.value[y].mail) {
                            $('<option />', { value: res2.tid, text: res2.tName }).appendTo(sOwners);
                            count++;
                        }
                        else if (count === 0 && email === res2.value[y].mail) {
                            $('<option />', { value: res2.tid, text: res2.tName }).appendTo(sMembers);
                        }
                        else if (count === 0 && email !== res2.value[y].mail) {
                            $('<option />', { value: res2.tid, text: res2.tName }).appendTo(sMembers);
                        }
                    }                                        
                });
        }    

        sOwners.appendTo($("#dTOwners"));
        sMembers.appendTo($("#dTMembers"));
      
    }



};


