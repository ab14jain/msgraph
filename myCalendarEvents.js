function getAccessToken(token){
    var authenticated = (window.location.href.split("#access_token=")[1]!=undefined);
    var token = "";
    if (!authenticated)
    {
        //authenticate user using OAuth
        requestToken();    
    }
    else
    {
        //there is a token in the URL. Let's retrieve from the query string
        token = window.location.href.split("#access_token=")[1].split("&")[0];  
    }

    var call = $.ajax({
        url: "https://outlook.office.com/api/v1.0/me/calendarview?startDateTime="+startString+"&endDateTime=" + endString,    
        type: "GET",
        dataType: "json",
        headers: {
            Accept: "application/json;odata.metadata=minimal;odata.streaming=true",
            'Authorization': "Bearer " + token
        }    
    });

    call.done(function (data, textStatus, jqXHR){

        var events = [];
        //loop through the returned events and push them
        //into the events array that will be displayed in the calendar.
        for (index in data.value)
        {
            events.push({
                        title: data.value[index].Subject,
                        start: data.value[index].Start
                    });
        }


        //callback(events);

    });

    call.fail(function (jqXHR,textStatus,errorThrown){
        alert("Error retrieving events: " + jqXHR.responseText);
    });    
}

function requestToken() { 
    // Change clientId and replyUrl to reflect your app's values 
    // found on the Configure tab in the Azure Management Portal
    // App ID is generated
    var clientId = '87979fdb-127f-42f0-a56b-95c1b82664cc';      
    var replyUrl    = 'https://jhomechef.sharepoint.com/sites/Northwind/SitePages/Home.aspx'; 
    var endpointUrl = 'https://outlook.office365.com/ews/odata/Me/Events';
    var resource = "https://outlook.office365.com/";        
    var authServer  = 'https://login.windows.net/common/oauth2/authorize?';  
    var responseType = 'token'; 

    var url = authServer + 
            "response_type=" + encodeURI(responseType) + "&" + 
            "client_id=" + encodeURI(clientId) + "&" + 
            "resource=" + encodeURI(resource) + "&" + 
            "redirect_uri=" + encodeURI(replyUrl); 

    //redirect user to the OAuth URL. The user will be redirected back to the page
    //unless some error occurs that prevents it
    window.location = url; 
}