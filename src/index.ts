import csvtojson from 'csvtojson'
import nodemailer from 'nodemailer';
import _ from 'underscore';

main().then(() => {});

let sender = ""
let position = "Events Officer"
let email = "docsoc@ic.ac.uk"
let password = process.env.DOCSOC_PASS

async function main() {
    if (process.argv.length == 5) {

        let eventName = process.argv[2];

        let responses = _.shuffle(await readCSV(process.argv[3]));
        let vouchers = await readCSV(process.argv[4]);
        if (vouchers.length > responses.length) {
            console.log("The following vouchers are UNUSED:");
            vouchers.slice(responses.length).forEach((voucher) => {
                console.log(voucher.code);
            });
        }

        var results = [];

        for (var i = 0; i < Math.min(vouchers.length, responses.length); i++) {
            results.push({
                email: responses[i].Email,
                name: responses[i].Name,
                voucher: vouchers[i].code,
            });
        }

        let mailer = nodemailer.createTransport({
    	    host: 'smtp.office365.com',
    	    port: 587,
    	    secure: false,
    	    auth: {
    		    user: email,
    		    pass: password,
    	    },
        });

        for (var i = 0; i < results.length; i++) {
            let result = results[i];
            await mailer.sendMail({
                from: email,
                to: result.email,
                subject: eventName + " - Deliveroo Voucher Winner",
                html: generateBody(eventName, result.name, result.voucher, sender, position)
            });

            console.log("Sent code " + result.voucher + " to " + result.name + " (" + result.email + ")");
        }
    } else {
        console.log("Usage: 'npm start <EVENT NAME> <RESPONSES> <VOUCHERS>");
    }
}

function readCSV(path: string) {
  return csvtojson().fromFile(path);
}

function generateBody(event: string, recipient: string, code: string, sender: string, position: string) {

    let template =
`<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset="iso-8859-1">
        <style type="text/css" style="display:none;"> P {margin-top:0;margin-bottom:0;} </style>
    </head>
    <body dir="ltr">
        <div style="font-family: Calibri, Helvetica, sans-serif; font-size: 11pt; color: rgb(0, 0, 0);">Hi [[NAME]]!</div>
        <div style="font-family: Calibri, Helvetica, sans-serif; font-size: 11pt; color: rgb(0, 0, 0);"><br></div>
        <div style="font-family: Calibri, Helvetica, sans-serif; font-size: 11pt; color: rgb(0, 0, 0);">Thank you for attending the [[EVENT]]! Hope you enjoyed it!</div>
        <div style="font-family: Calibri, Helvetica, sans-serif; font-size: 11pt; color: rgb(0, 0, 0);">You have won the Deliveroo voucher draw! Your £10 voucher code (valid for one year) is:</div>
        <div style="font-family: Calibri, Helvetica, sans-serif; font-size: 11pt; color: rgb(0, 0, 0);"><br></div>
        <div style="font-family: Calibri, Helvetica, sans-serif; font-size: 11pt; color: rgb(0, 0, 0);">[[CODE]]</div>
        <div style="font-family: Calibri, Helvetica, sans-serif; font-size: 11pt; color: rgb(0, 0, 0);"><br></div>
        <div id="Signature">
            <div>
                <div></div>
                <div id="divtagdefaultwrapper" dir="ltr" style="font-size:12pt; color:rgb(0,0,0); background-color:rgb(255,255,255); font-family:Calibri,Arial,Helvetica,sans-serif,EmojiFont,'Apple Color Emoji','Segoe UI Emoji',NotoColorEmoji,'Segoe UI Symbol','Android Emoji',EmojiSymbols">
                    <div style="font-family:Tahoma; font-size:13px">
                        <div style="font-family:Tahoma; font-size:13px">
                            <div>
                                <font size="2">
                                    <span style="font-size:10pt">
                                        <div style="color:rgb(33,33,33); font-size:15px; margin:0px">
                                            <span style="color:rgb(0,0,0); font-family:Calibri,sans-serif; font-size:12pt">
                                                <span style="color:rgb(33,33,33); font-family:Calibri,sans-serif; font-size:14.6667px">Kind regards,</span>
                                            </span>
                                        </div>
                                        <div style="color:rgb(33,33,33); font-size:15px; margin:0px">
                                            <b style="font-size:14pt; font-family:Calibri,sans-serif">[[SENDER]]</b><br>
                                        </div>
                                        <div style="color:rgb(33,33,33); font-size:15px; margin:0px">
                                            <font size="3" face="Times New Roman,serif">
                                                <span style="font-size:12pt">
                                                    <font size="2" face="Calibri,sans-serif">
                                                        <span style="font-size:11pt">DoCSoc [[POSITION]] 2020-2021</span>
                                                    </font>
                                                </span>
                                            </font>
                                        </div>
                                        <div style="color:rgb(33,33,33); font-size:15px; margin:0px">
                                            <font size="3" face="Times New Roman,serif">
                                                <span style="font-size:12pt">
                                                    <font size="2" face="Calibri,sans-serif">
                                                        <span style="font-size:11pt">Imperial College London</span>
                                                    </font>
                                                </span>
                                            </font>
                                            <b style="font-size:5pt; font-family:Calibri,sans-serif">&nbsp;<br><br></b>
                                        </div>
                                        <div style="color:rgb(33,33,33); font-size:15px; margin:0px">
                                            <font size="3" face="Times New Roman,serif">
                                                <span style="font-size:12pt">
                                                    <a href="http://docsoc.co.uk/" target="_blank">
                                                        <font size="2" face="Calibri,sans-serif" color="#1F4E79">
                                                            <span style="font-size:11pt">
                                                                <b>docsoc.co.uk</b>
                                                            </span>
                                                        </font>
                                                    </a>
                                                    <font size="2" face="Calibri,sans-serif">
                                                        <span style="font-size:11pt">&nbsp;|&nbsp;</span>
                                                    </font>
                                                    <a href="http://facebook.com/ICDoCSoc" target="_blank">
                                                        <font size="2" face="Calibri,sans-serif" color="#1F4E79">
                                                            <span style="font-size:11pt">
                                                                <b>facebook.com/ICDoCSoc</b>
                                                            </span>
                                                        </font>
                                                    </a>
                                                    <font size="2" face="Calibri,sans-serif">
                                                        <span style="font-size:11pt">&nbsp;|&nbsp;</span>
                                                    </font>
                                                    <a href="http://twitter.com/icdocsoc" target="_blank">
                                                        <font size="2" face="Calibri,sans-serif" color="#1F4E79">
                                                            <span style="font-size:11pt">
                                                                <b>@icdocsoc</b>
                                                            </span>
                                                        </font>
                                                    </a>
                                                </span>
                                            </font>
                                        </div>
                                    </span>
                                </font>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </body>
</html>`;

    let result = template
        .replace(/\[\[NAME\]\]/, recipient)
        .replace(/\[\[EVENT\]\]/, event)
        .replace(/\[\[CODE\]\]/, code)
        .replace(/\[\[SENDER\]\]/, sender)
        .replace(/\[\[POSITION\]\]/, position);

    return result;
}