if (typeof (TT) == "undefined")
    TT = {};

TT.Helpers = {
    callingAction: function (pluralEntityLogicalName, entityId, actionName, successCallback, errorCallback, isAsync, actionParameters) {
        debugger;
        var clientUrl = Xrm.Page.context.getClientUrl();
        var url = clientUrl + "/api/data/v8.2/" + pluralEntityLogicalName + "(" + entityId.replace("{", "").replace("}", "") + ")/Microsoft.Dynamics.CRM." + actionName;

        var req = new XMLHttpRequest();
        req.open("POST", url, isAsync);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.onreadystatechange = function () {
            if (req.readyState === 4) {
                if (req.status === 200) {
                    if (req.responseText !== null && req.responseText !== undefined && req.responseText != "") {
                        var resp = JSON.parse(req.responseText);
                        successCallback(JSON.parse(req.responseText));
                    }
                }
                else if (req.status === 204) {
                    successCallback();
                }
                else {
                    errorCallback(JSON.parse(req.response).error.message);
                }
            }
        };

        req.send(actionParameters !== null && actionParameters !== undefined ? JSON.stringify(actionParameters) : null);
    }
}

TT.RibbonCommands = {
    Execute: function () {
        Xrm.Page.ui.clearFormNotification("sucessNotify");
        Xrm.Page.ui.clearFormNotification("errorNotify");

        var entityId = Xrm.Page.data.entity.getId();
        var currentUserId = Xrm.Page.context.getUserId();
        var params = {
            "CurrentUser": {
                "@odata.type": "Microsoft.Dynamics.CRM.systemuser",
                "systemuserid": currentUserId.replace("{", "").replace("}", "")
                }
        };
        TT.Helpers.callingAction(
            "contacts",
            entityId,
            "new_SentEmails",
            function (args) {
                TT.RibbonCommands.SuccessCallback(args);
            },
            function (args) {
                TT.RibbonCommands.ErrorCallback(args)
            },
            true,
            params
            );
    },
    SuccessCallback: function (args) {
        Xrm.Page.ui.setFormNotification("E-mail send Successed!", "INFO", "sucessNotify");
    },
    ErrorCallback: function (args) {
        Xrm.Page.ui.setFormNotification("Error: " + args, "ERROR", "errorNotify");
    }
}