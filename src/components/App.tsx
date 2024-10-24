import * as React from "react";
import { Spinner, SpinnerType } from "office-ui-fabric-react";
import OfficeAddinMessageBar from "./OfficeAddinMessageBar";
import { signInO365, logoutFromO365 } from "../../utilities/office-apis-helpers";
import Frame1 from "./Frame1";
import Frame2 from "./Frame2";
import Frame3 from "./Frame3";
import ImmoMailScreen from "./ImmoMailScreen";
export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  authStatus?: string;
  errorMessage?: string;
  currentFrame: string;
  requestInput?: string;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      authStatus: "notLoggedIn",
      errorMessage: "",
      currentFrame: "default",
    };

    // Bind methods to the class instance
    this.boundSetState = this.setState.bind(this);
    this.setToken = this.setToken.bind(this);
    this.setUserName = this.setUserName.bind(this);
    this.displayError = this.displayError.bind(this);
    this.login = this.login.bind(this);
    this.logout = this.logout.bind(this);
    this.switchToFrame1 = this.switchToFrame1.bind(this);
    this.switchToFrame2 = this.switchToFrame2.bind(this);
    this.switchToFrame3 = this.switchToFrame3.bind(this);
  }

  /*
        Properties
    */

  // The access token is not part of state because React is concerned with the
  // UI and the token is not used to affect the UI in any way.
  accessToken: string;
  userName: string;

  /*
        Methods
    */

  boundSetState: () => {};

  setToken = (accesstoken: string) => {
    this.accessToken = accesstoken;
  };

  setUserName = (userName: string) => {
    this.userName = userName;
  };

  displayError = (error: string) => {
    this.setState({ errorMessage: error });
  };

  // Runs when the user clicks the X to close the message bar where
  // the error appears.
  errorDismissed = () => {
    this.setState({ errorMessage: "" });

    // If the error occurred during an "in process" phase (logging in),
    // the action didn't complete, so return the UI to the preceding state/view.
    this.setState((prevState) => {
      if (prevState.authStatus === "loginInProcess") {
        return { authStatus: "notLoggedIn" };
      }
      return null;
    });
  };

  login = async () => {
    await signInO365(
      this.boundSetState,
      this.setToken,
      this.setUserName,
      this.displayError
    );
  };

  logout = async () => {
    await logoutFromO365(
      this.boundSetState,
      this.setUserName,
      this.userName,
      this.displayError
    );
  };

  switchToFrame1 = () => {
    this.setState({ currentFrame: "Frame1" });
  };

  // Modify `switchToFrame2` to accept `requestInput`
  switchToFrame2 = (requestInput: string) => {
    this.setState({ currentFrame: "Frame2", requestInput });
  };

  switchToFrame3 = () => {
    this.setState({ currentFrame: "Frame3" });
  };

  render() {
    const { isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Spinner
          className="spinner"
          type={SpinnerType.large}
          label="Please sideload your add-in to see app body."
        />
      );
    }

    // Set the body of the page based on where the user is in the workflow.
    let body;

    if (this.state.currentFrame === "Frame1") {
      body = (
        <Frame1
          switchToFrame2={this.switchToFrame2}
          displayError={this.displayError}
          accessToken={this.accessToken}
        />
      );
    } else if (this.state.currentFrame === "Frame2") {
      body = (
        <Frame2
          switchToFrame3={this.switchToFrame3}
          accessToken={this.accessToken}
          requestInput={this.state.requestInput}
        />
      );
    } else if (this.state.currentFrame === "Frame3") {
      body = (
        <Frame3
          accessToken={this.accessToken}
          requestInput={this.state.requestInput}
        />
      );
    } else if (this.state.authStatus === "notLoggedIn") {
      body = <ImmoMailScreen login={this.login} />;
    } else if (this.state.authStatus === "loginInProcess") {
      body = (
        <Spinner
          className="spinner"
          type={SpinnerType.large}
          label="Bitte melden Sie sich im Popup-Fenster an."
        />
      );
    }

    return (
      <div>
        {this.state.errorMessage ? (
          <OfficeAddinMessageBar
            onDismiss={this.errorDismissed}
            message={this.state.errorMessage + " "}
          />
        ) : null}

        <div className="ms-welcome">
          {body}
          {/* Optionally, include a button to go back to home */}
          {/* <button onClick={this.switchToFrame1} style={{ position: "fixed", bottom: 0, width: "100%" }}>
            Go to Home
          </button> */}
        </div>
      </div>
    );
  }
}
