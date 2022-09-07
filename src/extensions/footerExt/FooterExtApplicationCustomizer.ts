import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";
import styles from "./FooterExt.module.scss";
import * as strings from "FooterExtApplicationCustomizerStrings";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { ISocial } from "./Models/ISocial";

export interface ISPLISTITEM {
  value: any[];
}
const LOG_SOURCE: string = "FooterExtApplicationCustomizer";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IFooterExtApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** Une class qui contients des actions à executé au lancement de l'extension */
export default class FooterExtApplicationCustomizer extends BaseApplicationCustomizer<IFooterExtApplicationCustomizerProperties> {
  /**
   * On doit déclarer une variable nommé BootomPlaceholder
   * Site sharepoint moderne  dispose 2 placeholders fixe où on mets nos footer et Header
   * cette variable doit définir le contenu de placeholder en bas de site sharepoint
   */
  private BootomPlaceHolder: PlaceholderContent | undefined;
  /**
      
     * Comme les extensions n'ont pas une propretypane pour configuré les listes on utilise des données static
      *     On définit  host url de la web application utilisé
     * le titre de a liste où cntient nos items Dans notre cas la listes des socials medias
     */
  
  private HostUrl: string = "https://intranetpprd.cpa-bank.dz/"; // prepre
  private ListName: string = "SocialLinks";
  private MenuList: string = "FooterData";
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    /**
     * On Affiche le footer définit dans la méthode renderPlaceholders
     */
    this.context.placeholderProvider.changedEvent.add(
      this,
      this.renderPlaceholders
    );
    return Promise.resolve();
  }
  /**
   * Cette méthode permet de récupérer les items d'une liste sharepoint ( social medias )
   * @param listTitle Paramètre represente le titre de la liste à utiliser
   * @returns  Retourner les éléments de la liste en paramètre de type array de Iscial contient les informations sur les social medias
   */
  private async getSocialMedias(listTitle: string): Promise<ISPLISTITEM> {
    // consummer API sharepoint
    return this.context.spHttpClient
      .get(
        `${this.HostUrl}/_api/Web/Lists/getbytitle('${listTitle}')/Items?$select=ID,Code,SocialIcon,SocialLink`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json(); // retourner la liste
      });
  }
  private async getMenuLinks(listTitle: string): Promise<ISPLISTITEM> {
    // consummer API sharepoint
    return this.context.spHttpClient
      .get(
        `${this.HostUrl}/_api/Web/Lists/getbytitle('${listTitle}')/Items?$select=ID,Title,Link&$orderby=Ordre&$top=6`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json(); // retourner la liste
      });
  }
  /**
   * C'est une méthode permet de définir le contenu de footer et le mettre en place dans le site et ensuite l'affiche
   */
  private async renderPlaceholders(): Promise<void> {
    // récupérer la listes des socials medias on appllant la méthode getListItems
    const SocialsList: ISocial[] = await (
      await this.getSocialMedias(this.ListName)
    ).value;
    const logo = SocialsList.filter((value) => {
      return value.Code == "logo";
    });
    const socialLinks = SocialsList.filter((value) => value.Code != "logo");
    const NavLinks: any[] = await (
      await this.getMenuLinks(this.MenuList)
    ).value;
    // --------------------------------------------------Footer---------------------------------------------

    // récupérer  bottomPlaceholder
    if (!this.BootomPlaceHolder) {
      // créer le   Footer
      this.BootomPlaceHolder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Bottom,
          { onDispose: this.onPlaceholderDispose }
        );
    }
    //tester si le footer est disponible
    if (this.BootomPlaceHolder) {
      // Mettre en place le contenu de footer
      if (this.BootomPlaceHolder.domElement) {
        this.BootomPlaceHolder.domElement.innerHTML = `
      <div
        class=${styles.footer}
      >
        <div class=${styles.container}>
          <div class=${styles.flex}>
            <div>
              <img src=${logo[0].SocialIcon.Url} class=${styles.icon} />
            </div>
            <div class="${styles.navbar}">
              ${
                NavLinks && NavLinks.length
                  ? NavLinks.map(
                      (link) =>
                        ` 
                  <ul class="${styles.ulstyle}">
                    <li style="float: left;"><a class="${styles.MenuLinks}" href="${link.Link}">${link.Title}</a><li>
                  </ul>
                `
                    ).join(" ")
                  : `<p >Copyright@ Crédit Populaire Algérie<p>`
              }
            </div>
            <div class=${styles.socialsContainer}>
              ${socialLinks
                .map(
                  (value) =>
                    `  <div>
                          <a target="_blank" href=${value.SocialLink.Url}>
                            <img src=${value.SocialIcon.Url} />
                          </a>
                        </div> `
                )
                .join(" ")}
            </div>
          </div>
        </div>
      </div>`;
      }
    }
  }
  /**
   * i called when placeholder disposed
   */
  private onPlaceholderDispose() {
    console.log("Placeholder disposed");
  }
}
