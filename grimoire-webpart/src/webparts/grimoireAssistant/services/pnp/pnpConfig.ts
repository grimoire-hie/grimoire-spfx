/**
 * PnPjs Configuration
 * Initializes and provides access to the SharePoint PnPjs instance
 */

import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/search";
import "@pnp/sp/site-users/web";
import type { WebPartContext } from "@microsoft/sp-webpart-base";
import { getContext, isInitialized, setContext } from "./pnpContext";

let _sp: SPFI | undefined;

/**
 * Initialize or retrieve the PnPjs SharePoint instance
 */
export const getSP = (context?: WebPartContext): SPFI => {
  if (context) {
    setContext(context);
    _sp = spfi().using(SPFx(context));
  }

  if (!_sp) {
    throw new Error(
      "PnPjs not initialized. Call getSP with WebPartContext first."
    );
  }

  return _sp;
};

export default getSP;
export { getContext, isInitialized };
