import * as React from 'react';
import styles from './WebcomFormularios.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';

export interface IWebcomFormulariosProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  formId: IPropertyPaneDropdownOption;
}

export const WebcomFormularios: React.FunctionComponent<IWebcomFormulariosProps> = (props: React.PropsWithChildren<IWebcomFormulariosProps>) => {

  return (
      <section className={`${styles.webcomFormularios} ${props.hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={props.isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Webcom Formularios</h2>
          <div>Seleccione el formulario que desea mostrar en la configuración</div>
          <div>{props.environmentMessage}</div>
          <div>{props.formId?.text}</div>
        </div>
      </section>
  );
};