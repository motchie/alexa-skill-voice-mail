'use strict';
var client = require("./lib/office365-rest-api-client");

var accessToken = "eyJ0eXAiOiJKV1QiLCJub25jZSI6IkFRQUJBQUFBQUFCSGg0a21TX2FLVDVYcmp6eFJBdEh6dUxnQXZBSlRycTNXSlJkWE1WRUlGMjVvX3ItZTR1NU1oX0E0UHdqZ19iYnR6SVJxXzFCc3BJczktb3haNUkxYU95MmFoMzh4cnBqMDVIeGlGWkh0d1NBQSIsImFsZyI6IlJTMjU2IiwieDV0IjoiejQ0d01kSHU4d0tzdW1yYmZhSzk4cXhzNVlJIiwia2lkIjoiejQ0d01kSHU4d0tzdW1yYmZhSzk4cXhzNVlJIn0.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9hMTk4MDAyYS03NTNjLTQwYzMtYTAyZi05YjY2YjkzYzM2MTUvIiwiaWF0IjoxNTE2MDMzMzUzLCJuYmYiOjE1MTYwMzMzNTMsImV4cCI6MTUxNjAzNzI1MywiYWNyIjoiMSIsImFpbyI6IlkyTmdZTGh0UCt0ZVE1cEVXWHpNMjl1TENveTltMXIxazF3bnhrMzZxZUgrSUp2eld6NEEiLCJhbXIiOlsicHdkIl0sImFwcF9kaXNwbGF5bmFtZSI6ImFsZXhhLXNraWxsLXZvaWNlLW1haWwiLCJhcHBpZCI6IjdlNDU0ZmJjLWVkMGItNDQyZi04YTRmLTc4ZDUxNzc0YmIxOCIsImFwcGlkYWNyIjoiMSIsImZhbWlseV9uYW1lIjoi5oyB55SwIiwiZ2l2ZW5fbmFtZSI6IuW-uSIsImlwYWRkciI6IjU4LjgwLjk0LjE3OCIsIm5hbWUiOiLmjIHnlLAg5b65Iiwib2lkIjoiY2NjZDIzMTQtMzdhMC00OTg0LTg1N2MtZmQ4OTllNTUyYTEyIiwicGxhdGYiOiIzIiwicHVpZCI6IjEwMDM3RkZFODQ0REM5NzQiLCJzY3AiOiJNYWlsLlJlYWQgVXNlci5SZWFkIiwic3ViIjoidHNPWXpYYnN2LUF2XzlqN1YtcHY4Wk9rWFI0NGx0SFFfdkFsdjVIWWhwWSIsInRpZCI6ImExOTgwMDJhLTc1M2MtNDBjMy1hMDJmLTliNjZiOTNjMzYxNSIsInVuaXF1ZV9uYW1lIjoibW90Y2hpZUBtb3RjaGllLmNvbSIsInVwbiI6Im1vdGNoaWVAbW90Y2hpZS5jb20iLCJ1dGkiOiJscllIQlBfWngwdUJqOUljcmdraUFBIiwidmVyIjoiMS4wIiwid2lkcyI6WyI2MmU5MDM5NC02OWY1LTQyMzctOTE5MC0wMTIxNzcxNDVlMTAiXX0.D1gRkVpPr2yt6RHdXFpaunpXmYGrkktgFY52KNpQEa_ut93EAPt8_edtYgpMIZaFu5Y70JlL1FKLYRerOk2DfRkb54B_WR45P7CLZmW0ujeF3ZAgfSppk45k72scjGEv1AfzKuVUiZ8f9gjLHUGoEp15oav28NKVEMWLLZICJ8jPdfANq75o4x_51RLuoMGMP5f1am76M3EW0vPxaOWyqafoAcAkuEsIVAZl3SuKTk3dEBpfadbL-Z7mcKane1MJfuCluo91c1RFbcsUP6xPfNo5gJI36GD-A6hgYaDk5dIGIxbky5_mEgdOTl5GjyQ99PMvn9Q68wC67tdsvRkaSg";
client.setAccessToken(accessToken);
var unReadMailCount;

client.UnReadMails()
    .then(
        (value) => {
            console.log(value); //unReadMailCount = value;
        }
    )
    .catch(
        (error) => { console.log(error); }
    );

// console.log(unReadMailCount);

// client.unReadMailCount()
//     .then(
//         (value) => {
//             console.log(value); //unReadMailCount = value;
//         }
//     )
//     .catch(
//         (error) => { console.log(error); }
//     );

// console.log(unReadMailCount);