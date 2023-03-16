import * as React from "react";
import styles from "./ServiceTicket.module.scss";
import { CircularProgress } from "@material-ui/core";
import {
  TextField,
  Label,
  ITextFieldStyles,
  Spinner,
  Dropdown,
  IDropdownStyles,
} from "@fluentui/react";
import { useState, useEffect } from "react";
import { useBoolean } from "@fluentui/react-hooks";
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";
import "../../../ExternalRef/css/Style.css";

interface IServiceTicket {
  Title: string;
  Description: string;
  ServiceTicketTypes: any;
}

interface IServiceTicketError {
  Title: string;
  Description: string;
  ServiceTicketTypes: string;
}

interface IserviceType {
  key: string;
  text: string;
}

let serviceType: IserviceType[] = [];

const MainServiceTicket = (props: any) => {
  /* Variable-decluration section start */
  let serviceObj: IServiceTicket = {
    Title: "",
    Description: "",
    ServiceTicketTypes: null,
  };

  let errService: IServiceTicket = {
    Title: "",
    Description: "",
    ServiceTicketTypes: null,
  };
  /* Variable-decluration section end */

  /* State-decluration section start */
  const [masterService, setMasterService] =
    useState<IServiceTicket>(serviceObj);
  const [errorService, setErrorService] = useState<IServiceTicket>(errService);
  const [errorAdminmsg, seterrorAdminmsg] = useState<string>();
  const [isLoader, setIsLoader] = useState<boolean>(true);
  const [isSpinner, setIsSpinner] = useState<boolean>(false);
  const [curUser, setCurUser] = useState("");
  const [successmsg, setsuccessmsg] = useState("");
  const [disablebtn, setdisablebtn] = useState(false);
  const [hideDialog, { toggle: toggleHideDialog }] = useBoolean(true);

  const modalPropsStyles = { main: { maxWidth: 450 } };
  const dialogContentProps = {
    type: DialogType.normal,
    title: "IT Service Ticket",
    subText: "Email our salesforce tech team at techteam@hosthealthcare.com.",
  };

  /* State-decluration section end */

  /* Style section start */
  const longTextBoxStyles: Partial<ITextFieldStyles> = {
    root: { width: "100%", margin: "5px 0px" },
    fieldGroup: {
      height: 80,
      border: "1px solid #00584d",
      borderRadius: "3px",
      ":after": { border: "1px solid #00584d" },
    },
    field: { fontSize: 14 },
  };

  const textBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: "100%",
      margin: "5px 10px 0px 0px",
    },
    fieldGroup: {
      border: "1px solid #00584d",
      borderRadius: "3px",
      ":after": { border: "1px solid #00584d" },
    },
    field: {
      fontSize: 14,
    },
  };

  const textBoxErrorStyles: Partial<ITextFieldStyles> = {
    root: { width: "100%", margin: "5px 10px 0px 0px" },
    fieldGroup: {
      border: "2px solid #f00",
      ":hover": { border: "2px solid #f00" },
      ":after": { border: "2px solid #f00" },
    },
    field: { fontSize: 14 },
  };

  const templateDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: "100%",
      ".ms-Dropdown-title": {
        border: "1px solid #00584d",

        boxShadow: "0px 3px 20px #888b8d1f",
        borderRadius: "3px",
        ":hover": {
          border: "1px solid #00584d",
        },
      },
      ".ms-Dropdown": {
        ":hover": {
          border: "none",
        },
        ":focus::after": {
          border: "none",
        },
      },
    },
  };

  const dropdownErrorStyles: Partial<IDropdownStyles> = {
    root: {
      width: "100%",
      ".ms-Dropdown-title": {
        border: "2px solid red",
        boxShadow: "0px 3px 20px #888b8d1f",
        background: "#fff",
        borderRadius: "3px",
        ":hover": {
          border: "2px solid red",
        },
      },
      ".ms-Dropdown": {
        ":hover": {
          border: "none",
        },
        ":focus::after": {
          border: "none",
        },
      },
    },
  };
  /* Style section end */

  /* Function section start */
  // get Dropdown value function
  const getDropdownValue = async () => {
    await props.sp.web.lists
      .getByTitle("IT Service Ticket")
      // .getByTitle("Service Ticket")IT Service Ticket
      .fields.getByInternalNameOrTitle("ServiceTicketTypes")
      .get()
      .then((response: any) => {
        serviceType=[];
        if (response.Choices.length > 0) {
          response.Choices.forEach((choice: string) => {
            serviceType.push({
              key: choice,
              text: choice,
            });
          });
          setIsLoader(false);
        } else {
          setIsLoader(false);
        }
      })
      .catch((error: any) => {
        getErrorFunction(error);
      });
  };

  // get error function section
  const getErrorFunction = (error: any) => {
    setdisablebtn(true);
    seterrorAdminmsg("Something went wrong. Please contact system admin.")
    setIsSpinner(false);
    setIsLoader(false);
    setTimeout(() => {
      seterrorAdminmsg("");
      setdisablebtn(false);
    }, 2000);
  };

  // validation function section
  const getvalidation = () => {
    let errValidation: IServiceTicketError = {
      Title: "",
      Description: "",
      ServiceTicketTypes: "",
    };
    if (!masterService.Title) {
      errValidation.Title = "Please enter name";
      setErrorService({ ...errValidation });
      setIsSpinner(false);
    } else if (!masterService.ServiceTicketTypes) {
      errValidation.ServiceTicketTypes = "Please select service ticket type";
      setErrorService({ ...errValidation });
      setIsSpinner(false);
    } else if (!masterService.Description) {
      errValidation.Description = "Please enter your description";
      setErrorService({ ...errValidation });
      setIsSpinner(false);
    } else {
      setErrorService({ ...errValidation });
      setdisablebtn(true);
      addRecord();
    }
  };

  // get record add function section
  const addRecord = async () => {
    await props.sp.web.lists
      .getByTitle("IT Service Ticket")
      // .getByTitle("Service Ticket")IT Service Ticket
      .items.add(masterService)
      .then((response: any) => {
        setMasterService({
          Title: "",
          Description: "",
          ServiceTicketTypes: null,
        });
        setIsSpinner(false);
        setsuccessmsg("Your service ticket was submitted succesfully");
        setTimeout(() => {
          setsuccessmsg("");
          setdisablebtn(false);
        }, 2000);
      })
      .catch((error: any) => {
        getErrorFunction(error);
      });
  };

  // useEffect function section
  useEffect(() => {
    setIsLoader(true);
    getDropdownValue();
    props.sp.web.currentUser.get().then((res) => {
      //masterService.Title = res.Title;
      //setMasterService({ ...masterService });
    });
  }, []);
  /* Function section end */

  return (
    <>
      {isLoader ? (
        <>
          <div
            style={{
              display: "flex",
              justifyContent: "center",
              alignItems: "center",
              height: "100vh",
            }}
          >
            <CircularProgress style={{ color: "blue" }} />
          </div>
        </>
      ) : (
        <>
          {/* Feedback section start */}
          <div className={styles.FeedBackBodySection}>
            {/* Lable section */}
            <div className={styles.Header}>IT Service Ticket</div>

            {/* Feedback Text Field section start */}
            <div>
              <div style={{ marginTop: 12 }}>
                <Label required className={styles.LabelSection}>
                  1. Your Name
                </Label>
                <TextField
                  styles={
                    errorService.Title && !disablebtn
                      ? textBoxErrorStyles
                      : textBoxStyles
                  }
                  placeholder="Enter your name"
                  value={masterService.Title}
                  onChange={(e: any, value: string) => {
                    masterService.Title = value;
                    setMasterService({ ...masterService });
                  }}
                />
              </div>
              <div style={{ marginTop: 12 }}>
                <Label required className={styles.LabelSection}>
                  2. Ticket Related To
                </Label>
                <Dropdown
                  className="TicketDropDown"
                  placeholder="Select your ticket type"
                  styles={
                    errorService.ServiceTicketTypes && !disablebtn
                      ? dropdownErrorStyles
                      : templateDropdownStyles
                  }
                  selectedKey={masterService.ServiceTicketTypes}
                  options={serviceType}
                  onChange={(e, option) => {
                    masterService.ServiceTicketTypes = option.key;
                    setMasterService({ ...masterService });

                    if (option.key == "Salesforce Related") {
                      toggleHideDialog();
                      setdisablebtn(true);
                    } else {
                      setdisablebtn(false);
                    }
                  }}
                />
              </div>
              <div style={{ marginTop: 12 }}>
                <Label required className={styles.LabelSection}>
                  3. Description
                </Label>
                <TextField
                  styles={
                    errorService.Description && !disablebtn
                      ? textBoxErrorStyles
                      : longTextBoxStyles
                  }
                  placeholder="Enter description"
                  multiline
                  value={masterService.Description}
                  onChange={(e: any, value: string) => {
                    masterService.Description = value;
                    setMasterService({ ...masterService });
                  }}
                />
              </div>
            </div>
            {/* Feedback Text Field section end */}

            {/* BTN section start */}
            <div>
              {!disablebtn ? (<>
                <div style={{ color: "red", fontWeight: "600" }}>
                  {errorService.Title ? `* ${errorService.Title}` : ""}
                  {errorService.Description
                    ? `* ${errorService.Description}`
                    : ""}
                  {errorService.ServiceTicketTypes
                    ? `* ${errorService.ServiceTicketTypes}`
                    : ""}
                </div>
                </>
              ) : (
                ""
              )}
              <div style={{ color: "green", fontWeight: "600" }}>
                {successmsg}
              </div>
              <div style={{ color: "red", fontWeight: "600" }}>
                {(!errorService.Title&&!errorService.Description&&!errorService.ServiceTicketTypes)&&errorAdminmsg?errorAdminmsg:""}
              </div>
              <button
                disabled={disablebtn}
                className={styles.btnSection}
                onClick={() => (setIsSpinner(true), getvalidation())}
              >
                {isSpinner ? <Spinner /> : "SUBMIT"}
              </button>
              <div className="dialogSection">
                <Dialog
                  styles={{
                    main: {
                      minHeight: "unset",
                      innerContent: {
                        display: "flex",
                        flexDirection: "column",
                      },
                    },
                  }}
                  hidden={hideDialog}
                  // onDismiss={toggleHideDialog}
                  dialogContentProps={dialogContentProps}
                >
                  <button
                    className={styles.dialogClose}
                    onClick={() => {
                      toggleHideDialog();
                      masterService.ServiceTicketTypes = "";
                      setMasterService({ ...masterService });
                      setdisablebtn(false);
                    }}
                  >
                    Close
                  </button>
                </Dialog>
              </div>
            </div>
            <div className="clsTextSubmit"><label style={{color:"#00584d"}}><b>Note</b>: Create an IT service ticket for Kam at Wendego. Salesforce-related questions should be directed to the Technology department at techteam@hosthealthcare.com.</label></div>
            {/* BTN section end */}
          </div>
          {/* Feedback section end */}
        </>
      )}
    </>
  );
};

export default MainServiceTicket;
