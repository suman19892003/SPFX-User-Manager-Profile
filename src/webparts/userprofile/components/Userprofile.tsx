import * as React from 'react';
import styles from './Userprofile.module.scss';
import { IUserprofileProps } from './IUserprofileProps';
import { escape, find } from '@microsoft/sp-lodash-subset';
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { sp } from "@pnp/sp/presets/all";

export default class Userprofile extends React.Component<IUserprofileProps, any> {

  constructor(props){
    super(props);
    this.state = {
      userId:"",
      //Code from Global Claims
      userDetailedInformation: [
        {
          PIN: "",
          OfficeNumber: "",
          SBU: "",
          Phone: "",
          Fax: "",
          Manager_Phone: "",
          DepartmentNumber: "",
          Employee_Email: "",
          Manager_Email: "",
          Employee_Name: "",
          OfficeName: "",
          DepartmentName: "",
          Job_Title: "",
          Job_Code: "",
          Manager_PIN: "",
          Manager_Name: "",
        },
      ]
    };
  }
  
private getUserDetails = () => {
  sp.profiles.getPropertiesFor(this.state.userId).then((profile) => {
    sp.profiles
      .getPropertiesFor(
        find(profile.UserProfileProperties, ["Key", "Manager"]).Value
      )
      .then((Managerprofile) => {
        this.setState({
          userDetailedInformation: [
            {
              PIN: find(profile.UserProfileProperties, [
                "Key",
                "enterprisePIN",
              ]).Value,
              Employee_Name: find(profile.UserProfileProperties, [
                "Key",
                "PreferredName",
              ]).Value,
              Employee_Email: find(profile.UserProfileProperties, [
                "Key",
                "UserName",
              ]).Value,
              OfficeName: find(profile.UserProfileProperties, [
                "Key",
                "Office",
              ]).Value,
              OfficeNumber: find(profile.UserProfileProperties, [
                "Key",
                "LMOfficeNumber",
              ]).Value,
              DepartmentName: find(profile.UserProfileProperties, [
                "Key",
                "Department",
              ]).Value,
              DepartmentNumber: find(profile.UserProfileProperties, [
                "Key",
                "DeptID",
              ]).Value,
              Job_Title: find(profile.UserProfileProperties, [
                "Key",
                "SPS-JobTitle",
              ]).Value,
              Job_Code: find(profile.UserProfileProperties, [
                "Key",
                "LMJobCode",
              ]).Value,
              Fax: find(profile.UserProfileProperties, ["Key", "Fax"]).Value,
              Phone: find(profile.UserProfileProperties, ["Key", "CellPhone"])
                .Value,
              Manager_PIN: find(Managerprofile.UserProfileProperties, [
                "Key",
                "enterprisePIN",
              ]).Value,
              Manager_Name: find(Managerprofile.UserProfileProperties, [
                "Key",
                "PreferredName",
              ]).Value,
              Manager_Email: find(Managerprofile.UserProfileProperties, [
                "Key",
                "UserName",
              ]).Value,
              Manager_Phone: find(Managerprofile.UserProfileProperties, [
                "Key",
                "CellPhone",
              ]).Value,
              SBU: "Commercial Insurance",
            },
          ],
          requestFor: "Global Claim Center",
        });
      });
  });
}

private getUserDetailsOnDemand = async () => {
  debugger;
  if (this.state.userId == "") {
    let currentUser = await sp.web.currentUser();
    this.setState({
      userId: currentUser.LoginName,
    });
  }
  console.log(this.state);
  debugger;
  this.getUserDetails();
};

public _getPeoplePickerItems = async (items: any[]) => {
  this.setState({ userId: items[0].loginName });
};

  public render(): React.ReactElement<any> {
    return (<>
      <div className={styles.container}>               
      <div className={styles.mainheading}>User Info Details</div>
          <ul className={styles.tabcontent}>
            <li className={styles.panel}>
              <div className={styles.formlabel}>User Name/Email:</div>
              <div className={styles.usernamesec}>
                <div className={styles.usernameinput}>
                  <PeoplePicker
                      context={this.props.context}
                      //context={this.context}
                      personSelectionLimit={1}
                      showtooltip={true}
                      required={true}
                      disabled={false}
                      onChange={this._getPeoplePickerItems}
                      showHiddenInUI={false}
                      ensureUser={true}
                      principalTypes={[PrincipalType.User]}
                      resolveDelay={1000}
                  />
                <span className={styles.errormsg}>User Last Name, First Name</span>
                </div>
                <div className={styles.usernamecontrols}>

                <div className={styles.btnCol}>
                    <PrimaryButton
                      className={styles.ylwBtn}
                      text="Get User Info"
                      onClick={() => this.getUserDetailsOnDemand()}
                    />
                  </div>
                </div>
              </div>
            </li>
          </ul>
        </div>
      <div className={styles.Section1}>
      <h4 className={styles.headingBgColor}> USER INFORMATION SECTION</h4>
      <div className={styles.formCard}>
        <div>
          <div className={styles.formRow}>
            <div className={styles.formCol50}>
              <div className={styles.dFlex}>
                <div className={styles.formTextlabel}>
                  <label>Employee PIN:</label>
                </div>
                <div className={styles.formTextValue}>
                  <label>
                    {this.state.userDetailedInformation[0].PIN}
                  </label>
                </div>
              </div>
            </div>
            <div className={styles.formCol50}>
              <div className={styles.dFlex}>
                <div className={styles.formTextlabel}>
                  <label>Employee Name:</label>
                </div>
                <div className={styles.formTextValue}>
                  <label>
                    {this.state.userDetailedInformation[0].Employee_Name}
                  </label>
                </div>
              </div>
            </div>
          </div>
          <div className={styles.formRow}>
            <div className={styles.formCol50}>
              <div className={styles.dFlex}>
                <div className={styles.formTextlabel}>
                  <label>Office Number:</label>
                </div>
                <div className={styles.formTextValue}>
                  <label>
                    {this.state.userDetailedInformation[0].OfficeNumber}
                  </label>
                </div>
              </div>
            </div>
            <div className={styles.formCol50}>
              <div className={styles.dFlex}>
                <div className={styles.formTextlabel}>
                  <label>Office Name:</label>
                </div>
                <div className={styles.formTextValue}>
                  <label>
                    {this.state.userDetailedInformation[0].OfficeName}
                  </label>
                </div>
              </div>
            </div>
          </div>
          <div className={styles.formRow}>
            <div className={styles.formCol50}>
              <div className={styles.dFlex}>
                <div className={styles.formTextlabel}>
                  <label>Department Number:</label>
                </div>
                <div className={styles.formTextValue}>
                  <label>
                    {
                      this.state.userDetailedInformation[0]
                        .DepartmentNumber
                    }
                  </label>
                </div>
              </div>
            </div>
            <div className={styles.formCol50}>
              <div className={styles.dFlex}>
                <div className={styles.formTextlabel}>
                  <label>Department Name:</label>
                </div>
                <div className={styles.formTextValue}>
                  <label></label>
                  {this.state.userDetailedInformation[0].DepartmentName}
                </div>
              </div>
            </div>
          </div>
          <div className={styles.formRow}>
            <div className={styles.formCol50}>
              <div className={styles.dFlex}>
                <div className={styles.formTextlabel}>
                  <label>Employee Email:</label>
                </div>
                <div className={styles.formTextValue}>
                  <label>
                    {this.state.userDetailedInformation[0].Employee_Email}
                  </label>
                </div>
              </div>
            </div>
            <div className={styles.formCol50}>
              <div className={styles.dFlex}>
                <div className={styles.formTextlabel}>
                  <label>Phone:</label>
                </div>
                <div className={styles.formTextValue}>
                  <label>
                    {this.state.userDetailedInformation[0].Phone}
                  </label>
                </div>
              </div>
            </div>
          </div>
          <div className={styles.formRow}>
            <div className={styles.formCol50}>
              <div className={styles.dFlex}>
                <div className={styles.formTextlabel}>
                  <label>Fax:</label>
                </div>
                <div className={styles.formTextValue}>
                  <label>
                    {this.state.userDetailedInformation[0].Fax}
                  </label>
                </div>
              </div>
            </div>
            <div className={styles.formCol50}>
              <div className={styles.dFlex}>
                <div className={styles.formTextlabel}>
                  <label>Job Code</label>
                </div>
                <div className={styles.formTextValue}>
                  <label>
                    {this.state.userDetailedInformation[0].Job_Code}
                  </label>
                </div>
              </div>
            </div>
          </div>
          <div className={styles.formRow}>
            <div className={styles.formCol50}>
              <div className={styles.dFlex}>
                <div className={styles.formTextlabel}>
                  <label>Job Title</label>
                </div>
                <div className={styles.formTextValue}>
                  <label>
                    {this.state.userDetailedInformation[0].Job_Title}
                  </label>
                </div>
              </div>
            </div>
            <div className={styles.formCol50}>
              <div className={styles.dFlex}>
                <div className={styles.formTextlabel}>
                  <label>SBU</label>
                </div>
                <div className={styles.formTextValue}>
                  <label>
                    {this.state.userDetailedInformation[0].SBU}
                  </label>
                </div>
              </div>
            </div>
          </div>
          <div className={styles.formRow}>
            <div className={styles.formCol50}>
              <div className={styles.dFlex}>
                <div className={styles.formTextlabel}>
                  <label>Manager Name</label>
                </div>
                <div className={styles.formTextValue}>
                  <label>
                    {this.state.userDetailedInformation[0].Manager_Name}
                  </label>
                </div>
              </div>
            </div>
            <div className={styles.formCol50}>
              <div className={styles.dFlex}>
                <div className={styles.formTextlabel}>
                  <label>Manager PIN</label>
                </div>
                <div className={styles.formTextValue}>
                  <label>
                    {this.state.userDetailedInformation[0].Manager_PIN}
                  </label>
                </div>
              </div>
            </div>
          </div>
          <div className={styles.formRow}>
            <div className={styles.formCol50}>
              <div className={styles.dFlex}>
                <div className={styles.formTextlabel}>
                  <label>Manager Email</label>
                </div>
                <div className={styles.formTextValue}>
                  <label>
                    {this.state.userDetailedInformation[0].Manager_Email}
                  </label>
                </div>
              </div>
            </div>
            <div className={styles.formCol50}>
              <div className={styles.dFlex}>
                <div className={styles.formTextlabel}>
                  <label>Manager Phone</label>
                </div>
                <div className={styles.formTextValue}>
                  <label>
                    {this.state.userDetailedInformation[0].Manager_Phone}
                  </label>
                </div>
              </div>
            </div>
          </div>
          </div>
          </div>
        </div>
        </>
    );
  }
}
