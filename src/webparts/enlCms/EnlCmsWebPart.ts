import { Version } from '@microsoft/sp-core-library';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import styles from './EnlCmsWebPart.module.scss';
// import * as strings from 'EnlCmsWebPartStrings';

export interface IEnlCmsWebPartProps {
  description: string;
}

export default class EnlCmsWebPart extends BaseClientSideWebPart<IEnlCmsWebPartProps> {

  protected onInit(): Promise<void> {

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.enlCms} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">

      <div>

        <h4>Learn more about SPFx development:</h4>
      </div>
    </section>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

}
