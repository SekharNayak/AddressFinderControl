import { IInputs, IOutputs } from "./generated/ManifestTypes";
import axios from "axios";


export class AddressFinder
  implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private _value: string;

  // PCF framework delegate which will be assigned to this object which would be called whenever any update happens.
  private _notifyOutputChanged: () => void;

  // label element created as part of this control
  private label: HTMLInputElement;

  // button element created as part of this control
  private button: HTMLButtonElement;

  // Reference to the control container HTMLDivElement
  // This element contains all elements of our custom control example
  private _container: HTMLDivElement;

  private _innerDiv: HTMLDivElement;

  private _form: HTMLFormElement;

  private _errorMessage: HTMLInputElement;

  private _context: ComponentFramework.Context<IInputs>;

  private _formContext : any ;

  /**
   * Empty constructor.
   */
  constructor() {}

  /**
   * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
   * Data-set values are not initialized here, use updateView.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
   * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
   * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
   * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
   */
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ) {
    this._context = context;
    // Creating the label for the control and setting the relevant values.
    this.label = document.createElement("input");
    this.label.setAttribute("type", "label");
    this.label.addEventListener("blur", this.onInputBlur.bind(this));
    //Create a button to increment the value by 1.
    this.button = document.createElement("button");

    // Get the localized string from localized string
    this.button.innerHTML = context.resources.getString(
      "TS_IncrementControl_ButtonLabel"
    );

    this.button.classList.add("SimpleIncrement_Button_Style");
    this._notifyOutputChanged = notifyOutputChanged;
    //this.button.addEventListener("click", (event) => { this._value = this._value + 1; this._notifyOutputChanged();});
    this.button.addEventListener("click", this.onButtonClick.bind(this));

    // Adding the label and button created to the container DIV.
    this._container = document.createElement("div");
    this._container.appendChild(this.label);
    this._container.appendChild(this.button);
    container.appendChild(this._container);
    
  }

  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   */
  public updateView(context: ComponentFramework.Context<IInputs>): void {
    // This method would rerender the control with the updated values after we call NotifyOutputChanged
    //set the value of the field control to the raw value from the configured field
    this._value = context.parameters.postcode.raw!;
    this.label.value = this._value != null ? this._value.toString() : "";

    if (context.parameters.postcode.error) {
      this.label.classList.add("SimpleIncrement_Input_Error_Style");
    } else {
      this.label.classList.remove("SimpleIncrement_Input_Error_Style");
    }
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
   */
  public getOutputs(): IOutputs {
    // custom code goes here - remove the line below and return the correct output
    let result: IOutputs = {
      postcode: this._value,
    };
    return result;
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    // Add code to cleanup control if necessary
  }

  /**
   * Button Event handler for the button created as part of this control
   * @param event
   */
  private onButtonClick(event: Event): void {
    this.removeMessage();
    const url = `https://addressfinderapi.azurewebsites.net/api/v1/address/search?postCode=${this._value}`;
    axios
      .get(url, {
        headers: {
          "x-api-key": `26b407e0-33f7-4b85-abae-7be57baedad5`,
        },
      })
      .then((response) => {
        const {
          msfsi_AddressStreet1: addressLine1,
          msfsi_AddressCity: city,
          msfsi_AddressCountryRegion: country,
          msfsi_AddressZIPPostalCode: postcode,
        } = response.data;

        let curentRecord: ComponentFramework.EntityReference = {
          id: (<any>this._context).page.entityId,
          name: (<any>this._context).page.entityTypeName,
        };


        if(!curentRecord.id){
          this.addMessage("Please create a record first ",true);
          return;
        }

        console.log(curentRecord);

        console.log(`entity id ${curentRecord.id.toString()}`);

        //update given customer entity with the address details 
        const address = {
          addressLine1,
          city,
          postcode,
          country
        }

        this.updateRecord(curentRecord.name , curentRecord.id.toString(), address );
      })
      .then(_ => {
        const id = (<any>this._context).page.entityId.toString();
        const name = (<any>this._context).page.entityTypeName;
        this._context.navigation.openForm({ entityName : name , 
          entityId : id });
      })
      .catch((error) => {
        console.log(error);
        this.addMessage(error.message , true);
      });
    
  }

  /**
   * Input Blur Event handler for the input created as part of this control
   * @param event
   */
  private onInputBlur(event: Event): void {
    console.log("On input blur called ");
    this._value = this.label.value;
    this._notifyOutputChanged();
  }


  private updateRecord(entityName : string , id :string , address : any ) : void {
    this._context.webAPI
    .updateRecord(entityName, id, {
      msfsi_addressline1: address.addressLine1,
      msfsi_addresszippostalcode: address.postcode,
      msfsi_addresscountry: address.country,
      msfsi_addresscity: address.city,
    })
    .then((data) => {
      console.log(data);
      this.addMessage("Address found . Please refresh the page for the changes to be reflected " , false);
    })
    .catch((error) => {
      console.log(error);
      this.addMessage(error, true);
    });
  }

  private removeMessage(): void {
    let divToBeRemoved = document.getElementById("palceholder");
    if (divToBeRemoved != null) {
      this._container.removeChild(divToBeRemoved);
    }
  }

  private addMessage(message: string , isError : boolean ): void {
    let placeholderDiv = document.createElement("div");
    placeholderDiv.setAttribute("id", "palceholder");
    this._errorMessage = document.createElement("input");
    this._errorMessage.setAttribute("type", "label");
    this._errorMessage.classList.add(isError ? "error_message_style" : "success_message_style");
    this._errorMessage.value = message;
    placeholderDiv.appendChild(this._errorMessage);
    this._container.appendChild(placeholderDiv);
  }

 
}
