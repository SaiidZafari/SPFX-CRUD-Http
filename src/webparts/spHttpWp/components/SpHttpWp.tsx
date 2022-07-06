/* eslint-disable @typescript-eslint/naming-convention */
/* eslint-disable dot-notation */
/* eslint-disable @typescript-eslint/no-floating-promises */
import * as React from 'react';
import styles from './SpHttpWp.module.scss';
import { ISpHttpWpProps } from './ISpHttpWpProps';
// import { ISpFXCrudProps } from "./ISpFXCrudProps";
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export default class SpHttpWp extends React.Component<ISpHttpWpProps, {}> {
  public render(): React.ReactElement<ISpHttpWpProps> {
    // const {
    //   description,    
    //   hasTeamsContext,
    // } = this.props;

    return (
      <div className={styles.SpFxCrud}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <p className={styles.description}>
                {escape(this.props.description)}
              </p>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Item Id:</div>
                <input type="text" id="Id" />
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Title</div>
                <input type="text" id="title" />
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>Full Name</div>
                <input type="text" id="name" />
              </div>
              <div className={styles.itemField}>
                <div className={styles.fieldLabel}>All Items:</div>
                <div id="allItems" />
              </div>
              <div className={styles.buttonSection}>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.createItem}>
                    Create
                  </span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.getItemById}>
                    Read
                  </span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.getAllItems}>
                    Read All
                  </span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.updateItem}>
                    Update
                  </span>
                </div>
                <div className={styles.button}>
                  <span className={styles.label} onClick={this.deleteItem}>
                    Delete
                  </span>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
 
  
// Create Item
  private createItem = (): void => {
    const body: string = JSON.stringify({
      'Title': document.getElementById("title")['value'],
      'name': document.getElementById("name")['value']
    });
    this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Employees')/items`,
      SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: body
    })
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          response.json().then((responseJSON) => {
            console.log(responseJSON);
            alert(`Item created successfully with ID: ${responseJSON.ID}`);
          });
        } else {
          response.json().then((responseJSON) => {
            console.log(responseJSON);
            alert(`Something went wrong! Check the error in the browser console.`);
          });
        }
      }).catch(error => {
        console.log(error);
      });
  }
 
  
// Get Item by ID
  private getItemById = (): void => {
    const id: number = document.getElementById('Id')['value'];
    if (id > 0) {
      this.props.context.spHttpClient.get(this.props.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('Employees')/items(${id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            response.json().then((responseJSON) => {
              console.log(responseJSON);
              document.getElementById('title')['value'] = responseJSON.Title;
              document.getElementById('name')['value'] = responseJSON.name;
            });
          } else {
            response.json().then((responseJSON) => {
              console.log(responseJSON);
              alert(`Something went wrong! Check the error in the browser console.`);
            });
          }
        }).catch(error => {
          console.log(error);
        });
    }
    else {
      alert(`Please enter a valid item id.`);
    }
  }
 
  
// Get all items
  private getAllItems = (): void => {
    this.props.context.spHttpClient.get(this.props.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('Employees')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      })
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          response.json().then((responseJSON) => {
            // eslint-disable-next-line @typescript-eslint/typedef
            let html = `<table><tr><th>ID</th><th>Title</th><th>Full Name</th></tr>`;
            responseJSON.value.map((item, index) => {
              html += `<tr><td>${item.ID}</td><td>${item.Title}</td><td>${item.name}</td></li>`;
            });
            html += `</table>`;
            document.getElementById("allItems").innerHTML = html;
            console.log(responseJSON);
          });
        } else {
          response.json().then((responseJSON) => {
            console.log(responseJSON);
            alert(`Something went wrong! Check the error in the browser console.`);
          });
        }
      }).catch(error => {
        console.log(error);
      });
  }
 
  
// Update Item
  private updateItem = (): void => {
    const id: number = document.getElementById('Id')['value'];
    const body: string = JSON.stringify({
      'Title': document.getElementById("title")['value'],
      'name': document.getElementById("name")['value']
    });
    if (id > 0) {
      this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Employees')/items(${id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: body
        })
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            alert(`Item with ID: ${id} updated successfully!`);
          } else {
            response.json().then((responseJSON) => {
              console.log(responseJSON);
              alert(`Something went wrong! Check the error in the browser console.`);
            });
          }
        }).catch(error => {
          console.log(error);
        });
    }
    else {
      alert(`Please enter a valid item id.`);
    }
  }
 
  
// Delete Item
  private deleteItem = (): void => {
    const id: number = parseInt(document.getElementById('Id')['value']);
    if (id > 0) {
      this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Employees')/items(${id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE'
          }
        })
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            alert(`Item ID: ${id} deleted successfully!`);
          }
          else {
            alert(`Something went wrong!`);
            console.log(response.json());
          }
        });
    }
    else {
      alert(`Please enter a valid item id.`);
    }
  }
}
     