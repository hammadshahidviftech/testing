import * as React from 'react';
import "@pnp/polyfill-ie11";
import { sp, ItemAddResult, ItemUpdateResult } from "@pnp/sp";
import { IMetadataupdateProps } from './IMetadataupdateProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  TextField,
  DefaultButton,
  Dropdown, DropdownMenuItemType, IDropdownOption, IDropdownStyles
} from 'office-ui-fabric-react';
export interface IMetadataupdateState {
  test: string;
  Category: any;
  ReviewNeeded: any;
  ClientShareable: any;
  Document_x0020_Status: any;
  Document_x0020_Type: any;
  CategoryOptions: any;
  DocTypeOptions: any;
  ID: any;
  DocumentName:any;
  DocumentType:any;
}
import * as jquery from 'jquery';
import { SPHttpClient, SPHttpClientResponse, IHttpClientOptions, SPHttpClientConfiguration, ISPHttpClientOptions } from '@microsoft/sp-http';
var FullPath = window.location.href;
var arrayOfParts = FullPath.split('/');
const SharepointSiteUrl = arrayOfParts.slice(0, 5).join("/");
const BASEURL = SharepointSiteUrl;
const LIST_UpdateMetaLibrary = 'Documents';

export default class Metadataupdate extends React.Component<IMetadataupdateProps, IMetadataupdateState> {
  private DocumentStatusoptions: any[] = [];
  private DocumentShareableoptions: any[] = [];
  private DocumentReviewNeeded: any[] = [];
  private Catagoryoptions: any[] = [];

  constructor(props) {
    super(props);
    sp.setup({
      sp: {
        headers: {
          Accept: "application/json;odata=verbose",
        },
        baseUrl: BASEURL
      },
    });

    this.state = {
      test: '',
      Category: [],
      ReviewNeeded: '',
      ClientShareable: '',
      Document_x0020_Status: '',
      Document_x0020_Type: [],
      CategoryOptions: [],
      DocTypeOptions: [],
      ID: 0,
      DocumentName:'',
      DocumentType:''
    };


  }
  componentDidMount() {
    console.log('in in');
    this._openFormUsingLink();
    this._getDocumentStatus();
    this._getDocumentShareable();
    //this._getReviewNeeded();
    this._getCategory();
    this._getDocType();
  }
  //  onChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
  //   if (item) {
  //     setSelectedKeys(
  //       item.selected ? [...selectedKeys, item.key as string] : selectedKeys.filter(key => key !== item.key),
  //     );
  //   }
  // };
  public getCatagoryChange(e) {
    let selectedcatagory = this.state.Category;
    const checkedValue = e.target.value;
    const checked = e.target.checked;
    if (checked) {
      selectedcatagory.push(checkedValue)
    }
    else {
      var x = selectedcatagory.indexOf(checkedValue);
      selectedcatagory.splice(x, 1);
    }
    this.setState({ Category: selectedcatagory })
  }
  public getDocTypeChange(e) {
    let selectedcatagory = this.state.Document_x0020_Type;
    const checkedValue = e.target.value;
    const checked = e.target.checked;
    if (checked) {
      selectedcatagory.push(checkedValue)
    }
    else {
      var x = selectedcatagory.indexOf(checkedValue);
      selectedcatagory.splice(x, 1);
    }
    this.setState({ Document_x0020_Type: selectedcatagory })
  }

  public render(): React.ReactElement<IMetadataupdateProps> {
    return (
      <div >
        <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
          <TextField maxLength={255} label="Document Name" name="text"
          readOnly={true}
          value={this.state.DocumentName.trim()}
            //onChanged={(event): void => {

            // }} 
            autoComplete="off"
            id="documentname" />
        </div>
        <label style={{fontSize:'14px',fontWeight:600}}>Category</label><br></br>
        <div className="row" style={{ display: 'inline-flex' }}>
          
          {this.state.CategoryOptions.map((item, index) => {
            let Catagory = this.state.Category.filter((user) => {
              return user.trim() == item.text.trim()
            });
            let IsChecked = Catagory.length > 0 ? true : false;
            return <div className="ms-Grid-col ms-sm3 ms-md3 ms-lg3" style={{ marginTop: '10px',marginRight:'15px'}}>

              <input style={{  height: '20px', width: '20px' }} type="checkbox"
                onChange={(event): void => {
                  this.getCatagoryChange(event);
                }}
                value={item.text}
                checked={IsChecked}
              ></input>
              <label>{item.text}</label>
            </div>
          })}
        </div>
        
        <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
          <Dropdown
            label='Client Shareable'
            onChanged={(event): void => {
              this.setState({ ClientShareable: event.text })
            }}
            placeholder={this.state.ClientShareable}
            id="clientshareable"
            className="validateinput"
            options={
              this.DocumentShareableoptions
            }
          />
        </div>
        <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
          <Dropdown
            label='Document Status'
            onChanged={(event): void => {
              this.setState({ Document_x0020_Status: event.text })
            }}
            placeholder={this.state.Document_x0020_Status}
            id="documentstatus"
            className="validateinput"
            options={
              this.DocumentStatusoptions
            }
          />
        </div>
        {/* <label style={{fontSize:'14x',fontWeight:600}}>Document Type</label><br></br> */}
        <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
          {/* <div className="row" style={{ display: 'inline-flex' }}>
           
            {this.state.DocTypeOptions.map((item, index) => {
              let Document_x0020_Type = this.state.Document_x0020_Type.filter((type) => {
                return type.trim() == item.text.trim()
              });
              let IsChecked = Document_x0020_Type.length > 0 ? true : false;
              return <div className="ms-Grid-col ms-sm3 ms-md3 ms-lg3" style={{ marginTop: '10px',marginRight:'15px'}}>
                
                <input style={{ height: '20px', width: '20px' }} type="checkbox"
                  onChange={(event): void => {
                    this.getDocTypeChange(event);
                  }}
                  value={item.text}
                  checked={IsChecked}
                ></input>
                <label>{item.text}</label>
              </div>
            })}
          </div> */}
          <Dropdown
            label='Document Type'
            onChanged={(event): void => {
              this.setState({ DocumentType: event.text })
            }}
            placeholder={this.state.DocumentType}
            id="DocumentType"
            className="DocumentType"
            options={
              this.state.DocTypeOptions
            }
          />
          
          </div>
        {/* <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
          <Dropdown
            label='Review Needed'
            onChanged={(event): void => {
              this.setState({ ReviewNeeded: event.text })
            }}
            placeholder={this.state.ReviewNeeded}
            id="reviewneeded"
            className="validateinput"
            options={
              this.DocumentReviewNeeded
            }
          />
        </div> */}
        <div style={{ marginTop: "10px" }}>
          <DefaultButton text="Update" style={{ float: "right" }}
            onClick={() => this._UpdateDocumentMetaData()}
          />
        </div>
      </div >
    );
  }
  private _openFormUsingLink(): void {

    let itemId = getUrlParameter("itemid");

    if (itemId !== false && itemId !== "" && itemId !== undefined) {
      itemId = isNaN(Number(itemId)) ? "" : Number(itemId);
      let url: string = SharepointSiteUrl + "/_api/web/lists/getByTitle('Documents')/items(" + itemId + ")?$select=*,File/ServerRelativeUrl&$expand=File";
      jquery.getJSON(url, {
        format: "json"
      })
        .done(responseJSON => {
          console.log("Get Document Data", responseJSON)
          if (responseJSON !== null && responseJSON !== "" && responseJSON !== undefined) {
            this.setState({
              ID: itemId,
              DocumentName:replaceNullsByEmpty(responseJSON.File.ServerRelativeUrl),
              Category: responseJSON.Category_x0020_1,
              ReviewNeeded: responseJSON.Review_x0020_Needed,
              ClientShareable: responseJSON.Client_x0020_Shareable,
              Document_x0020_Status: responseJSON.Document_x0020_Status,
              DocumentType: responseJSON.Document_x0020_Type
            })
          }
        }).fail(error => {

          console.log('Error' + error);
        });
    }
    else {

    }
  }
  private _getDocumentStatus(): void {
    let test = [];
    let url: string = SharepointSiteUrl + "/_api/web/lists/getByTitle('Documents')/fields?$filter=EntityPropertyName eq 'Document_x0020_Status'";
    jquery.getJSON(url, {
      format: "json"
    })
      .done(responseJSON => {
        for (let i = 0; i < responseJSON.value[0].Choices.length; i++) {
          this.DocumentStatusoptions.push({
            key: responseJSON.value[0].Choices[i],
            text: responseJSON.value[0].Choices[i]
          })

        }
      }).fail(error => {

        console.log('Error' + error);
      });

  }
  private _getCategory(): void {
    let test = [];
    let url: string = SharepointSiteUrl + "/_api/web/lists/getByTitle('Documents')/fields?$filter=EntityPropertyName eq 'Category_x0020_1'";
    jquery.getJSON(url, {
      format: "json"
    })
      .done(responseJSON => {
        console.log("Category_x0020_1", responseJSON)
        for (let i = 0; i < responseJSON.value[0].Choices.length; i++) {
          test.push({
            text: responseJSON.value[0].Choices[i].trim(),
          })

        }
        this.setState({ CategoryOptions: test })
      }).fail(error => {

        console.log('Error' + error);
      });

  }
  private _getDocType(): void {
    let test = [];
    let url: string = SharepointSiteUrl + "/_api/web/lists/getByTitle('Documents')/fields?$filter=EntityPropertyName eq 'Document_x0020_Type'";
    jquery.getJSON(url, {
      format: "json"
    })
      .done(responseJSON => {
        for (let i = 0; i < responseJSON.value[0].Choices.length; i++) {
          test.push({
            text: responseJSON.value[0].Choices[i].trim(),
            key: responseJSON.value[0].Choices[i].trim(),
          })

        }
        this.setState({ DocTypeOptions: test })
      }).fail(error => {

        console.log('Error' + error);
      });

  }
  // private _getReviewNeeded(): void {
  //   let test = [];
  //   let url: string = SharepointSiteUrl + "/_api/web/lists/getByTitle('Documents')/fields?$filter=EntityPropertyName eq 'Review_x0020_Needed'";
  //   jquery.getJSON(url, {
  //     format: "json"
  //   })
  //     .done(responseJSON => {
  //       for (let i = 0; i < responseJSON.value[0].Choices.length; i++) {
  //         this.DocumentReviewNeeded.push({
  //           key: responseJSON.value[0].Choices[i],
  //           text: responseJSON.value[0].Choices[i]
  //         })

  //       }
  //     }).fail(error => {

  //       console.log('Error' + error);
  //     });

  // }
  private _getDocumentShareable(): void {
    let test = [];
    let url: string = SharepointSiteUrl + "/_api/web/lists/getByTitle('Documents')/fields?$filter=EntityPropertyName eq 'Client_x0020_Shareable'";
    jquery.getJSON(url, {
      format: "json"
    })
      .done(responseJSON => {
        for (let i = 0; i < responseJSON.value[0].Choices.length; i++) {
          this.DocumentShareableoptions.push({
            key: responseJSON.value[0].Choices[i],
            text: responseJSON.value[0].Choices[i]
          })

        }
      }).fail(error => {

        console.log('Error' + error);
      });

  }
  public _UpdateDocumentMetaData() {
    try {
      let val = []
      let id = this.state.ID;
      sp.web.lists.getByTitle(LIST_UpdateMetaLibrary).items.getById(id).update({
        Category_x0020_1: { results: this.state.Category },
       // Review_x0020_Needed: this.state.ReviewNeeded,
        Client_x0020_Shareable: this.state.ClientShareable,
        Document_x0020_Status: this.state.Document_x0020_Status,
         Document_x0020_Type:this.state.DocumentType
      }).then((iar: ItemUpdateResult) => {
        console.log("Updated Item");
        alert("MetaData Updated")
        window.location.href =(SharepointSiteUrl)
      });

    } catch (error) {
      console.log("Error while updating Item: ", error);
    }
  }
}
function getUrlParameter(sParam: string): any {
  let sPageURL: string = window.location.search.substring(1),
    sURLVariables: string[] = sPageURL.split('&'),
    sParameterName: string[],
    i: number;
  for (i = 0; i < sURLVariables.length; i++) {
    sParameterName = sURLVariables[i].split('=');
    if (sParameterName[0].toLowerCase() === sParam) {
      return (sParameterName[1] === undefined || replaceNullsByEmptyString(sParameterName[1]) === "") ? false : decodeURIComponent(sParameterName[1]);
    }
  }
}
function replaceNullsByEmptyString(value) {
  return (value == null ) ? "" : value
}
function replaceNullsByEmpty(value) {
  return (value == null || value == undefined || value == "undefined") ? "" : value
}
