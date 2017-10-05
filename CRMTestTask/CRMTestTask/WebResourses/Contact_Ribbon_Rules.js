if (typeof (TT) == "undefined")
    TT = {};

TT.RibbonRules = {
    SendEmailDisplay: function () {
        var adminRole = "Системный администратор";
        var customiserRole = "Настройщик системы";
        var currentUserRoles = SVB.EnableRules.GetCurrentUserRoles();
        var isHasRole = false;
        for (var i = 0; i < currentUserRoles.length; i++) {
            if (currentUserRoles[i].name.toLowerCase() == secretaryRole.toLowerCase()) {
                isHasRole = true;
                break;
            }
        }

        return isHasRole;
    },
    GetCurrentUserRoles: function () {
        /// <summary>
        /// Function returns current user roles objects.
        /// </summary>

        var userRoleIds = Xrm.Page.context.getUserRoles();
        var userRoles = null;
        var filters = "";

        for (var i = 0; i < userRoleIds.length; i++) {
            if (filters == "") {
                filters += "&$filter=" + "roleid eq " + userRoleIds[i].replace("{", "").replace("}", "");
            }
            else {
                filters += " or roleid eq " + userRoleIds[i].replace("{", "").replace("}", "");
            }
        }

        TT.RibbonRules.RetriveEntitiesHelper(
            "roles",
            "?$select=roleid,name" + filters,
            function (retrievedRoles) {
                userRoles = retrievedRoles.value;
            },
            function (error) {
                alert(error);
            },
            false,
            false
        );

        return userRoles;
    },

    RetriveEntitiesHelper: function (entityPluralName, queryString, successCallback, errorCallback, isAsync, getFormattedValues) {
        ///<summary>
        /// Sends request to retrieve a record.
        ///</summary>
        ///<param name="entityPluralName" type="String">
        /// The Plural Name of the Entity type record to retrieve.
        /// For an Account record, use "accounts"
        ///</param>
        ///<param name="queryString" type="String">
        /// A String representing the Query Option to control which attributes will be returned 
        /// and/or which related records are also returned.
        ///</param>
        ///<param name="successCallback" type="Function">
        /// The function that will be passed through and be called by a successful response. 
        /// This function must accept the returned record as a parameter.
        /// </param>
        ///<param name="errorCallback" type="Function">
        /// The function that will be passed through and be called by a failed response. 
        /// This function must accept an Error object as a parameter.
        /// </param>
        ///<param name="isAsync" type="Boolean">
        /// A Boolean representing if the method should run asynchronously or synchronously
        /// true means asynchronously. false means synchronously
        ///</param>

        var req = new XMLHttpRequest();
        req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/" + entityPluralName + queryString, isAsync);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");

        if (getFormattedValues === true) {
            req.setRequestHeader("Prefer", "odata.include-annotations=\"OData.Community.Display.V1.FormattedValue\"");
        }

        req.onreadystatechange = function () {
            if (req.readyState === 4) {
                if (req.status === 200) {
                    successCallback(JSON.parse(req.responseText));
                }
                else {
                    errorCallback(JSON.parse(req.response).error.message);
                }
            }
        };

        req.send();
    }
}
