<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title>ASP Demo</title>
    </head>
    <body>
        <!-- 
            The below example demonstrates how to make an AJAX call to a SQL DB
            and return the data for a user using classic ASP
        -->
        <script>
            function showUser(str){
                if(str == ""){
                    document.getEelementById("users").innerHTML="";
                    return;
                }
                else if(window.XMLHttpRequest){
                    xmlhttp = new XMLHttpRequest();
                }
                xmlhttp.onreadystatechange = function(){
                    if(this.readyState == 4 && this.status == 200){
                        document.getElementById("users").innerHTML = this.responseText;
                    }
                }
                xmlhttp.open("GET", "GetUser.asp?q=" + str, true);
                xmlhttp.send();
            }
        </script>

        <form>
            <select name="users" onchange="showUser(this.value)">
                <option value="">Select a user:</option>
                <option value="Aaron">Aaron</option>
                <option value="Sarah">Sarah</option>
                <option value="Bluford">Bluford</option>
                <option value="Anne">Anne</option>
                <option value="Adam">Adam</option>
            </select>
        </form>
        <br>

        <div id="users"></div>
        
        <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.2/jquery.min.js"></script>
        <script type = "text/javascript"src = "https://ajax.googleapis.com/ajax/libs/jqueryui/1.11.3/jquery-ui.min.js"></script>
        <script src="Script.js"></script>
    </body>
</html>
