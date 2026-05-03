// ========================================
// SharePoint REST API Layer (Generic)
// ========================================
// Reusable SharePoint REST API module.
// No build step. All functions exposed via window.* globals.
// Auth: SharePoint cookie session (credentials: 'include') + X-RequestDigest for writes.
//
// Site URL Resolution Priority:
//   1. window._spRestSiteUrl  (manual override)
//   2. _spPageContextInfo.webAbsoluteUrl  (SharePoint runtime context)
//   3. window.location.origin  (fallback)

function _getSiteUrl() {
  if (window._spRestSiteUrl) return window._spRestSiteUrl;
  if (window._spPageContextInfo && window._spPageContextInfo.webAbsoluteUrl) {
    return window._spPageContextInfo.webAbsoluteUrl;
  }
  return window.location.origin;
}

window.__spDigestCache = null;
window.__spListTypeCache = window.__spListTypeCache || {};

async function getRequestDigest() {
  try {
    if (window.__spDigestCache) return window.__spDigestCache;
    const resp = await fetch(`${_getSiteUrl()}/_api/contextinfo`, {
      method: "POST",
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
      },
      credentials: "include",
    });
    if (!resp.ok) throw new Error(`contextinfo failed: ${resp.status}`);
    const data = await resp.json();
    window.__spDigestCache = data.d.GetContextWebInformation.FormDigestValue;
    // Cache expires after 25 minutes (SP default is 30min)
    setTimeout(() => { window.__spDigestCache = null; }, 25 * 60 * 1000);
    return window.__spDigestCache;
  } catch (e) {
    console.error("[SP-REST] getRequestDigest error:", e);
    return null;
  }
}

async function getListItemType(listTitle) {
  try {
    if (window.__spListTypeCache[listTitle]) return window.__spListTypeCache[listTitle];
    const resp = await fetch(
      `${_getSiteUrl()}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')?$select=ListItemEntityTypeFullName`,
      {
        method: "GET",
        headers: { Accept: "application/json;odata=verbose" },
        credentials: "include",
      }
    );
    if (!resp.ok) throw new Error(`getListItemType failed: ${resp.status}`);
    const data = await resp.json();
    window.__spListTypeCache[listTitle] = data.d.ListItemEntityTypeFullName;
    return window.__spListTypeCache[listTitle];
  } catch (e) {
    console.error("[SP-REST] getListItemType error:", e);
    return null;
  }
}

async function fetchListItems(listTitle, filter, select, orderby, top, expand) {
  try {
    const _filter = filter || "";
    const _select = select || "*";
    const _orderby = orderby || "";
    const _top = top || 500;
    const _expand = expand || "";
    let url = `${_getSiteUrl()}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$top=${_top}`;
    if (_filter) url += `&$filter=${encodeURIComponent(_filter)}`;
    if (_select) url += `&$select=${encodeURIComponent(_select)}`;
    if (_expand) url += `&$expand=${encodeURIComponent(_expand)}`;
    if (_orderby) url += `&$orderby=${encodeURIComponent(_orderby)}`;

    const results = [];
    let nextUrl = url;
    while (nextUrl) {
      const resp = await fetch(nextUrl, {
        method: "GET",
        headers: { Accept: "application/json;odata=verbose" },
        credentials: "include",
      });
      if (!resp.ok) throw new Error(`fetchListItems failed: ${resp.status}`);
      const data = await resp.json();
      results.push(...(data.d.results || []));
      nextUrl = data.d.__next || null;
    }
    return results;
  } catch (e) {
    console.error(`[SP-REST] fetchListItems(${listTitle}) error:`, e);
    return [];
  }
}

async function createSharePointListItem(listTitle, item) {
  try {
    const digest = await getRequestDigest();
    if (!digest) throw new Error("No request digest");
    const itemType = await getListItemType(listTitle);
    if (!itemType) throw new Error("Could not get list item type");
    const resp = await fetch(
      `${_getSiteUrl()}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items`,
      {
        method: "POST",
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": digest,
        },
        credentials: "include",
        body: JSON.stringify({ __metadata: { type: itemType }, ...item }),
      }
    );
    if (!resp.ok) {
      const errText = await resp.text();
      throw new Error(`Create failed: ${resp.status} - ${errText}`);
    }
    const data = await resp.json();
    return data.d;
  } catch (e) {
    console.error(`[SP-REST] createSharePointListItem(${listTitle}) error:`, e);
    return null;
  }
}

async function updateSharePointListItem(listTitle, id, item) {
  try {
    const digest = await getRequestDigest();
    if (!digest) throw new Error("No request digest");
    const itemType = await getListItemType(listTitle);
    if (!itemType) throw new Error("Could not get list item type");
    const resp = await fetch(
      `${_getSiteUrl()}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${id})`,
      {
        method: "POST",
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": digest,
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
        },
        credentials: "include",
        body: JSON.stringify({ __metadata: { type: itemType }, ...item }),
      }
    );
    if (!resp.ok) {
      const errText = await resp.text();
      throw new Error(`Update failed: ${resp.status} - ${errText}`);
    }
    return true;
  } catch (e) {
    console.error(`[SP-REST] updateSharePointListItem(${listTitle}, ${id}) error:`, e);
    return false;
  }
}

async function deleteSharePointListItem(listTitle, id) {
  try {
    const digest = await getRequestDigest();
    if (!digest) throw new Error("No request digest");
    const resp = await fetch(
      `${_getSiteUrl()}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${id})`,
      {
        method: "POST",
        headers: {
          Accept: "application/json;odata=verbose",
          "X-RequestDigest": digest,
          "IF-MATCH": "*",
          "X-HTTP-Method": "DELETE",
        },
        credentials: "include",
      }
    );
    if (!resp.ok) {
      const errText = await resp.text();
      throw new Error(`Delete failed: ${resp.status} - ${errText}`);
    }
    return true;
  } catch (e) {
    console.error(`[SP-REST] deleteSharePointListItem(${listTitle}, ${id}) error:`, e);
    return false;
  }
}

async function getUserProfile(loginName) {
  try {
    if (!loginName) return null;
    // PeoplePicker keys can come back in three forms:
    //   "i:0#.f|membership|user@domain"   (full claims string)
    //   "user@domain"                     (UPN / email only)
    //   "DOMAIN\\user"                    (sAM)
    // The PeopleManager endpoint expects the full claims form, encoded.
    let claims = loginName;
    if (!/^i:0#\.f\|/i.test(claims)) claims = "i:0#.f|membership|" + claims;
    const encoded = encodeURIComponent(claims).replace(/'/g, "%27%27");
    const resp = await fetch(
      `${_getSiteUrl()}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor('${encoded}')`,
      {
        method: "GET",
        headers: { Accept: "application/json;odata=verbose" },
        credentials: "include",
      }
    );
    if (!resp.ok) return null;
    const data = await resp.json();
    const props = (data.d && data.d.UserProfileProperties && data.d.UserProfileProperties.results) || [];
    const out = {};
    for (const p of props) {
      if (p && p.Key) out[p.Key] = p.Value || "";
    }
    return out;
  } catch (e) {
    return null;
  }
}

async function searchADUsers(searchTerm) {
  try {
    const digest = await getRequestDigest();
    if (!digest) throw new Error("No request digest");
    const resp = await fetch(
      `${_getSiteUrl()}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`,
      {
        method: "POST",
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": digest,
        },
        credentials: "include",
        body: JSON.stringify({
          queryParams: {
            __metadata: { type: "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters" },
            AllowEmailAddresses: true,
            AllowMultipleEntities: false,
            MaximumEntitySuggestions: 20,
            PrincipalSource: 15,
            PrincipalType: 1,
            QueryString: searchTerm,
          },
        }),
      }
    );
    if (!resp.ok) throw new Error(`User search failed: ${resp.status}`);
    const data = await resp.json();
    const results = JSON.parse(data.d.ClientPeoplePickerSearchUser) || [];

    // Enrich each result with WorkPhone/CellPhone/Department/Title from UserProfile.
    // PeoplePicker often returns those fields empty; UserProfile fills them in.
    // Failures degrade silently - the picker result alone is still returned.
    await Promise.all(
      results.map(async (r) => {
        const loginName = (r && (r.Key || r.Description)) || "";
        if (!loginName) return;
        const profile = await getUserProfile(loginName);
        if (!profile) return;
        r.EntityData = r.EntityData || {};
        if (!r.EntityData.WorkPhone   && profile.WorkPhone)       r.EntityData.WorkPhone   = profile.WorkPhone;
        if (!r.EntityData.MobilePhone && profile.CellPhone)       r.EntityData.MobilePhone = profile.CellPhone;
        if (!r.EntityData.Department  && profile.Department)      r.EntityData.Department  = profile.Department;
        if (!r.EntityData.Title       && profile["SPS-JobTitle"]) r.EntityData.Title       = profile["SPS-JobTitle"];
        if (!r.EntityData.Email       && profile.WorkEmail)       r.EntityData.Email       = profile.WorkEmail;
      })
    );

    return results;
  } catch (e) {
    console.error("[SP-REST] searchADUsers error:", e);
    return [];
  }
}

// Resolve a login claim or email -> SharePoint site user Id.
// SP Person fields require the integer Id (e.g. {OwnerId: 42}); EnsureUser
// accepts both claim strings (i:0#.w|domain\user) and email addresses.
async function ensureUserGetId(loginOrEmail) {
  if (!loginOrEmail) return null;
  try {
    const digest = await getRequestDigest();
    if (!digest) throw new Error("No request digest");
    const resp = await fetch(`${_getSiteUrl()}/_api/web/EnsureUser`, {
      method: "POST",
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest,
      },
      credentials: "include",
      body: JSON.stringify({ logonName: loginOrEmail }),
    });
    if (!resp.ok) throw new Error(`EnsureUser failed: ${resp.status}`);
    const data = await resp.json();
    return (data && data.d && data.d.Id) || null;
  } catch (e) {
    console.error(`[SP-REST] ensureUserGetId(${loginOrEmail}) error:`, e);
    return null;
  }
}

// Generic audit log helper. Pass the audit log list title as the first parameter.
async function createAuditLogEntry(auditListTitle, recordId, actionType, actorName, actorEmail, previousStatus, newStatus, details) {
  try {
    return await createSharePointListItem(auditListTitle, {
      Title: `${actionType} - Record ${recordId}`,
      RecordId: String(recordId),
      ActionType: actionType,
      ActorName: actorName || "",
      ActorEmail: actorEmail || "",
      PreviousStatus: previousStatus || "",
      NewStatus: newStatus || "",
      ActionDateTime: new Date().toISOString(),
      Details: details || "",
    });
  } catch (e) {
    console.error("[SP-REST] createAuditLogEntry error:", e);
    return null;
  }
}

// Expose to global scope (no build step)
window.getRequestDigest = getRequestDigest;
window.getListItemType = getListItemType;
window.fetchListItems = fetchListItems;
window.createSharePointListItem = createSharePointListItem;
window.updateSharePointListItem = updateSharePointListItem;
window.deleteSharePointListItem = deleteSharePointListItem;
window.searchADUsers = searchADUsers;
window.getUserProfile = getUserProfile;
window.ensureUserGetId = ensureUserGetId;
window.createAuditLogEntry = createAuditLogEntry;
