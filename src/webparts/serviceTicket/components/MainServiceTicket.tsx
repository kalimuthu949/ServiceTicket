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
  const [isLoader, setIsLoader] = useState<boolean>(true);
  const [isSpinner, setIsSpinner] = useState<boolean>(false);
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
        padding: "5px 10px",
        border: "1px solid #00584d",
        height: "40px",
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
        padding: "5px 10px",
        border: "2px solid red",
        height: "40px",
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
      .getByTitle("Service Ticket")
      .fields.getByInternalNameOrTitle("ServiceTicketTypes")
      .get()
      .then((response: any) => {
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
    console.log(error);
  };

  // validation function section
  const getvalidation = () => {
    let errValidation: IServiceTicketError = {
      Title: "",
      Description: "",
      ServiceTicketTypes: "",
    };
    if (!masterService.Title) {
      errValidation.Title = "Please enter your name";
      setErrorService({ ...errValidation });
      setIsSpinner(false);
    } else if (!masterService.Description) {
      errValidation.Description = "Please enter your description";
      setErrorService({ ...errValidation });
      setIsSpinner(false);
    } else if (!masterService.ServiceTicketTypes) {
      errValidation.ServiceTicketTypes = "Please select service ticket type";
      setErrorService({ ...errValidation });
      setIsSpinner(false);
    } else {
      setErrorService({ ...errValidation });
      addRecord();
    }
  };

  // get record add function section
  const addRecord = async () => {
    await props.sp.web.lists
      .getByTitle("Service Ticket")
      .items.add(masterService)
      .then((response: any) => {
        setMasterService({
          Title: "",
          Description: "",
          ServiceTicketTypes: null,
        });
        setIsSpinner(false);
      })
      .catch((error: any) => {
        getErrorFunction(error);
      });
  };

  // useEffect function section
  useEffect(() => {
    setIsLoader(true);
    getDropdownValue();
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
              <div>
                <Label required className={styles.LabelSection}>
                  1.Your Name
                </Label>
                <TextField
                  styles={
                    errorService.Title ? textBoxErrorStyles : textBoxStyles
                  }
                  placeholder="Enter your answer"
                  value={masterService.Title}
                  onChange={(e: any, value: string) => {
                    masterService.Title = value;
                    setMasterService({ ...masterService });
                  }}
                />
              </div>
              <div>
                <Label required className={styles.LabelSection}>
                  2.Describe the Problem
                </Label>
                <TextField
                  styles={
                    errorService.Description
                      ? textBoxErrorStyles
                      : longTextBoxStyles
                  }
                  placeholder="Enter your answer"
                  multiline
                  value={masterService.Description}
                  onChange={(e: any, value: string) => {
                    masterService.Description = value;
                    setMasterService({ ...masterService });
                  }}
                />
              </div>
              <div>
                <Label required className={styles.LabelSection}>
                  3.Ticket Related To
                </Label>
                <Dropdown
                  placeholder="Select your answer"
                  styles={
                    errorService.ServiceTicketTypes
                      ? dropdownErrorStyles
                      : templateDropdownStyles
                  }
                  selectedKey={masterService.ServiceTicketTypes}
                  options={serviceType}
                  onChange={(e, option) => {
                    masterService.ServiceTicketTypes = option.key;
                    setMasterService({ ...masterService });
                  }}
                />
              </div>
            </div>
            {/* Feedback Text Field section end */}

            {/* BTN section start */}
            <div>
              <div style={{ color: "red", fontWeight: "600" }}>
                {errorService.Title ? `* ${errorService.Title}` : ""}
                {errorService.Description
                  ? `* ${errorService.Description}`
                  : ""}
                {errorService.ServiceTicketTypes
                  ? `* ${errorService.ServiceTicketTypes}`
                  : ""}
              </div>
              <button
                disabled={isSpinner}
                className={styles.btnSection}
                onClick={() => (setIsSpinner(true), getvalidation())}
              >
                {isSpinner ? <Spinner /> : "SUBMIT"}
              </button>
            </div>
            {/* BTN section end */}
          </div>
          {/* Feedback section end */}
        </>
      )}
    </>
  );
};

export default MainServiceTicket;
