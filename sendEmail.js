const getEnvironmentVariableValue = (variableName) => {
    const entityName = "environmentvariablevalue";
    const fetchXml = `?fetchXml=
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1' no-lock='true'>
              <entity name='${entityName}'>
                <attribute name='value' />
                <link-entity name='environmentvariabledefinition' from='environmentvariabledefinitionid' to='environmentvariabledefinitionid' link-type='inner' alias='definition'>
                  <attribute name='defaultvalue' />
                  <filter type='and'>
                    <condition attribute='schemaname' operator='eq' value='${variableName}' />
                  </filter>
                </link-entity>
              </entity>
            </fetch>`;
    return new Promise(
        (resolve, reject) => {
            Xrm.WebApi.retrieveMultipleRecords(entityName, encodeURI(fetchXml)).then(
                (data) => {
                    if (data.entities.length > 0) {
                        if (data.entities[0].value) {
                            resolve(data.entities[0].value);
                        }
                        else if (data.entities[0]["definition.defaultvalue"]) {
                            console.log("Get by default value!");
                            resolve(data.entities[0]["definition.defaultvalue"]);
                        }
                        else {
                            reject({ message: `Environment variable called ${variableName} not defined!` });
                        }
                    }
                    else {
                        reject({ message: `Environment variable called ${variableName} not found!` });
                    }
                })
                .catch((error) => {
                    reject(error);
                });
        });
};

const getEmailTemplate = (entity, entityId, templateId) => {
    return new Promise(
        (resolve, reject) => {
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
            const variableName = "cr6c5_url_flow_send_email";
            getEnvironmentVariableValue(variableName)
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
                            console.log(data);
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
                })
                .catch((error) => {
                    reject(error);
                });
        });
};

const getContactEmailById = (id) => {
    const entityName = "contact";
    const fetchXml = `?fetchXml=
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1' no-lock='true'>
              <entity name='${entityName}'>
                <attribute name='emailaddress1' alias='email' />
                <filter type='and'>
                  <condition attribute='contactid' operator='eq' value='${id}'/>
                  <condition attribute='emailaddress1' operator='not-null' />
                </filter>
              </entity>
            </fetch>`;
    return new Promise(
        (resolve, reject) => {
            Xrm.WebApi.retrieveMultipleRecords(entityName, encodeURI(fetchXml)).then(
                (data) => {
                    if (data.entities.length > 0) {
                        resolve(data.entities[0].email);
                    }
                    else {
                        reject({ message: `The contact id ${id}, does not have an email!` });
                    }
                })
                .catch((error) => {
                    reject(error);
                });
        });
}

const getUserEmail = () => {
    const id = Xrm.Utility.getGlobalContext().userSettings.userId;
    const entityName = "systemuser";
    const fetchXml = `?fetchXml=
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' top='1' no-lock='true'>
              <entity name='${entityName}'>
                <attribute name='internalemailaddress' alias='email' />
                <filter type='and'>
                  <condition attribute='systemuserid' operator='eq' value='${id}'/>
                  <condition attribute='internalemailaddress' operator='not-null' />
                </filter>
              </entity>
            </fetch>`;
    return new Promise(
        (resolve, reject) => {
            Xrm.WebApi.retrieveMultipleRecords(entityName, encodeURI(fetchXml)).then(
                (data) => {
                    if (data.entities.length > 0) {
                        resolve(data.entities[0].email);
                    }
                    else {
                        reject({ message: `The user id ${id}, does not have an email!` });
                    }
                })
                .catch((error) => {
                    reject(error);
                });
        });
}

const sendEmailToContact = (id) => {
    getContactEmailById(id)
        .then((email) => {
            getUserEmail()
            .then((userEmail) => {
                getEmailTemplate("contact", id, "4a062d33-a6e5-1046-8195-ae46fc8f5a30")
                .then((data) => {
                    const body = {
                        to: email,
                        cc: userEmail,
                        subject: data.value[0].subject,
                        html: data.value[0].description
                    };
                    sendEmailByPowerAutomate(body)
                        .then((data) => {
                            console.log(data.isSuccess);
                        })
                        .catch((error) => {
                            console.log(`ERROR when sending email from Power Automate: ${error.message}`);
                        });
                })
                .catch((error) => {
                    console.log(`ERROR getting email template: ${error.message}`);
                });
            })
            .catch((error) => {
                console.log(`ERROR getting user email: ${error.message}`);
            });
        })
        .catch((error) => {
            console.log(`ERROR getting contact email: ${error.message}`);
        });
};

sendEmailToContact("9fd4a450-cb0c-ea11-a813-000d3a1b1223");
