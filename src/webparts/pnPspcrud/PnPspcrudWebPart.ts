import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PnPspcrudWebPart.module.scss';
import * as strings from 'PnPspcrudWebPartStrings';
import * as pnp from 'sp-pnp-js';
import { IPnPspcrudWebPartProps } from './IPnPspcrudWebPartProps';

		export interface ISPList {
      ID: string;
      Title: string;
      Experience: string;
      Location: string;
	} 

export default class PnPspCrudWebPart extends BaseClientSideWebPart<IPnPspcrudWebPartProps> {

  private AddEventListener(): void {
    document.getElementById('AddItem').addEventListener('click',()=>this.AddItem());
    document.getElementById('UpdateItem').addEventListener('click',()=>this.UpdateItem());
    document.getElementById('DeleteItem').addEventListener('click',()=>this.DeleteItem());
  }

  private _getListData(): Promise<ISPList[]> {
    
    return pnp.sp.web.lists.getByTitle("CadastroPessoa").items.get().then((response) => {      
       return response;
     });        
   }

  private getListData(): void {
   
    this._getListData()
      .then((response) => {
        this._renderList(response);
      });
}

  private _renderList(items: ISPList[]): void {
  let html: string = '<table class="TFtable" border=1 width=100% style="border-collapse: collapse;">';
  html += `<th>ID</th><th>Nome</th><th>Linguagem</th><th>Local</th>`;
  items.forEach((item: ISPList) => {
    html += `
         <tr>
        <td>${item.ID}</td>
        <td>${item.Title}</td>
        <td>${item.Experience}</td>
        <td>${item.Location}</td>
        </tr>
        `; 
  });
  html += `</table>`;
  const listContainer: Element = this.domElement.querySelector('#spGetListItems');
  listContainer.innerHTML = html;
}

  public render(): void {
  this.domElement.innerHTML = `    
        <div class="parentContainer" style="background-color: lightgrey">
    <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
    <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
        <span class="ms-font-xl ms-fontColor-white" style="font-size:28px">Cadastro de pessoas</span>
        <p class="ms-font-l ms-fontColor-white" style="text-align: left">SharePoint List CRUD using PnP JS and SPFx</p>
    </div>
    </div>
    <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
    <div style="background-color:Black;color:white;text-align: center;font-weight: bold;font-size:18px;">Detalhes</div>
    
    </div>
    <div style="background-color: lightgrey" >
    <form >
        <br>
        <div data-role="header">
          <h3>Adicionar item na lista</h3>
        </div>
        <div data-role="main" class="ui-content">
          <div >
              <input id="Title"  placeholder="Nome"    />
              <input id="Experience"  placeholder="Linguagem"  />
              <input id="Location"  placeholder="Local"    />
          </div>
          <div></br></div>
          <div >
              <button id="AddItem"  type="submit" >Adicionar</button>
          </div>
        </div>
        <div data-role="header">
          <h3>Atualizar/Deleta item na lista</h3>
        </div>
        <div data-role="main" class="ui-content">
          <div >
              <input id="EmployeeId"   placeholder="ID"  />
          </div>
          <div></br></div>
          <div >
              <button id="UpdateItem" type="submit" >Atualizar</button>
              <button id="DeleteItem"  type="submit" >Deletar</button>
          </div>
        </div>
    </form>
    </div>
    <br>
    <div style="background-color: lightgrey" id="spGetListItems" />
    </div>
    `;
    this.getListData();
    this.AddEventListener();
}

  private AddItem() {
    pnp.sp.web.lists.getByTitle('CadastroPessoa').items.add({
    Title : document.getElementById('Title')["value"],
    Experience : document.getElementById('Experience')["value"],
    Location:document.getElementById('Location')["value"]
   });
    alert("Registro salvo : "+ document.getElementById('Title')["value"] + " Adicionado !");  
}

  private UpdateItem() {
    var id = document.getElementById('EmployeeId')["value"];
    pnp.sp.web.lists.getByTitle("CadastroPessoa").items.getById(id).update({
      Title : document.getElementById('Title')["value"],
    Experience : document.getElementById('Experience')["value"],
    Location:document.getElementById('Location')["value"]
   });
    alert("Registro salvo : "+ document.getElementById('Title')["value"] + " Atualizado !");
 }
 
  private DeleteItem()
 {  
    pnp.sp.web.lists.getByTitle("CadastroPessoa").items.getById(document.getElementById('EmployeeId')["value"]).delete();
    alert("Registro : "+ document.getElementById('EmployeeId')["value"] + " Apagado !");
 }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
}

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
