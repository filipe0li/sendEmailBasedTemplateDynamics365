const getEnvironmentVariableValue = (variableName) => {
    return new Promise(
        (resolve, reject) => {
            const STATECODE_ACTIVE = 0;
            const query = `?$select=defaultvalue&$expand=environmentvariabledefinition_environmentvariablevalue($select=value;$filter=(statecode eq ${STATECODE_ACTIVE}))&$filter=(statecode eq ${STATECODE_ACTIVE} and displayname eq '${variableName}') and (environmentvariabledefinition_environmentvariablevalue/any(o1:(o1/statecode eq ${STATECODE_ACTIVE})))&$top=1`;
            Xrm.WebApi.retrieveMultipleRecords("environmentvariabledefinition", query).then(
                (data) => {
                    if (data.entities.length > 0) {
                        if (data.entities[0].environmentvariabledefinition_environmentvariablevalue[0].value) {
                            resolve(data.entities[0].environmentvariabledefinition_environmentvariablevalue[0].value);
                        }
                        else if (data.entities[0].defaultvalue) {
                            console.log("Get by default value!");
                            resolve(data.entities[0].defaultvalue);
                        }
                        else {
                            reject(`ERROR: Environment variable called ${variableName} not defined!`);
                        }
                    }
                    else {
                        reject(`ERROR: Environment variable called ${variableName} not found!`);
                    }
                })
                .catch((error) => {
                    reject(error);
                });
        });
};

const getEmailTemplate = (entity, entityId) => {
    return new Promise(
        (resolve, reject) => {
            const templateId = "4a062d33-a6e5-1046-8195-ae46fc8f5a30";
            const body = {
                TemplateId: templateId,
                ObjectType: entity,
                ObjectId: entityId
            };
            const url = `${Xrm.Utility.getGlobalContext().getClientUrl()}/api/data/v9.2/InstantiateTemplate`;
            const params = {
                method: "POST",
                headers: {
                    "Content-Type": "application/json; charset=utf-8"
                },
                body: JSON.stringify(body)
            };
            fetch(url, params)
                .then(response => response.json())
                .then(data => {
                    if (data.error) {
                        reject(data.error);
                    }
                    else {
                        resolve(data);
                    }
                })
                .catch((error) => {
                    reject(error);
                });
        });
};

const sendEmailByPowerAutomate = (body) => {
    /*
    const body = {
        to: "",
        cc: "",
        subject: "",
        body: html
    };
    */
    return new Promise(
        (resolve, reject) => {
            const sendEmailEnvironmentName = "url_flow_send_email";
            getEnvironmentVariableValue(sendEmailEnvironmentName)
                .then((url) => {
                    const params = {
                        method: "POST",
                        headers: {
                            "Content-Type": "application/json; charset=utf-8"
                        },
                        body: JSON.stringify(body)
                    };
                    fetch(url, params)
                        .then(response => response.json())
                        .then(data => {
                            resolve(data);
                        })
                        .catch((error) => {
                            reject(error);
                        });
                })
                .catch((error) => {
                    reject(error);
                });
        });
};

const sendEmailBasedTemplate = (entity, entityId) => {
    getEmailTemplate(entity, entityId)
        .then((data) => {
            const body = {
                to: "test@test.com",
                cc: "",
                subject: data.value[0].subject,
                html: data.value[0].description
            };
            return sendEmailByPowerAutomate(body);
        })
        .then(
            (isSucessful) => {
                console.log(isSucessful);
            }
        )
        .catch((error) => {
            console.log(error.message);
        });
};

sendEmailBasedTemplate("contact", "9fd4a450-cb0c-ea11-a813-000d3a1b1223");
