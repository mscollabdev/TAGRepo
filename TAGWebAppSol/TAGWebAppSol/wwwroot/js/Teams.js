window.addEventListener(
    "load",
    () => {
        init();
    },
    false,
);

let client;
let client1;
let client2;

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


    $('#boot-multiselect-demo').multiselect({
        nonSelectedText: 'Select Teams!',
        buttonWidth: 250,
        enableFiltering: true
    });

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
        for (var a = 0; a < teamIDs.length; a++) {
            var teamID = [];
            teamID = teamIDs[a].split('|');
            request
                .getTeamsOwners(teamID[0], teamID[1])                
                .then((res2) => {
                    for (var j = 0; j < res2.value.length; j++) {
                        var y = "'" + j + "'";
                        y = y.replace(/['"]+/g, '');

                        if (email === res2.value[y].mail) {
                            $('#boot-multiselect-demo').append('<option value=' + res2.tid + '>' + res2.tName + '</option>');
                        }
                    }
                
                });
        }
        
    }
   

};


