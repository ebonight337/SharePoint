import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './AccordionWebPart.module.scss';
import * as strings from 'AccordionWebPartStrings';

// jQuery関連ファイル
import * as jQuery from 'jquery';
import 'jqueryui';
// アコーディオンのテンプレート
// import MyAccordionTemplate from './MyAccordionTemplate';
// 外部CSSの読み込み
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IAccordionWebPartProps {
  description: string,
  listTitle: string;
}
// リストモデル
export interface ISPLists {
  value: ISPList[];
}

// 環境によって変更が必要！！ ////////////////////////////////////
export interface ISPList {
  Title: string;
  Url: string;
  DispName: string;
}
////////////////////////////////////////////////////////////////

// HTTPClient読み込み
import {
  SPHttpClient,
  SPHttpClientResponse,
} from '@microsoft/sp-http';
import { List } from 'lodash';



export default class AccordionWebPart extends BaseClientSideWebPart<IAccordionWebPartProps> {
  // カスタムメソッド /////////////////////////////////////////////////////////////

  //  jQueryUI スタイルを読み込み
  public constructor() {
    super();
    SPComponentLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
  }
  // リスト情報取得
  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('${this.properties.listTitle}')/Items`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then(jsonResponse => {
        console.log(jsonResponse.value);
        return jsonResponse;
      }) as Promise<ISPLists>;
  }
  // spListContainerIDにHTMLを入れ込む
  private _renderList(items: ISPList[]): void {
    let html: string = '';
    let count: number = 0;
    html += `<div class="accordion">`
    items.forEach((item: ISPList) => {
      console.log(item.Title + " " + item.DispName);
      // もし表示名が登録されていなければカテゴリと判断
      if(item.DispName == null){
        console.log(count);
        if(count == 0){
          html += `
          <h3>${item.Title}</h3>
          <div>`
          count ++;
        }else{
          html += `
          </div>
          <h3>${item.Title}</h3>
          <div>`
        }
      }else{
        html += `
          <span class="${ styles.accordionItem }">
            <a href="${item.Url}">${item.DispName}</a></br>
          </span>`
      }
    });
    html += `</div></div>`;
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
    
    // jQuery UI アコーディオンのカスタムプロパティ
    const accordionOptions: JQueryUI.AccordionOptions = {
      animate: true,
      collapsible: false,
      icons: {
        header: 'ui-icon-circle-arrow-e',
        activeHeader: 'ui-icon-circle-arrow-s'
      }
    };

    // 初期化
    jQuery('.accordion', this.domElement).accordion(accordionOptions);
  }

  private _renderListAsync(): void {
    // リストデータ取得
    this._getListData()
    .then((response) => {
      // リストレンダリング
      this._renderList(response.value);
    });
  }
  ////////////////////////////////////////////////////////////////////////////////
  
  public render(): void {
    this.domElement.innerHTML = `
      <div id="spListContainer" />
    `;
      this._renderListAsync();
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
                }),
                PropertyPaneTextField('listTitle', {
                  label: 'listTitle'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
