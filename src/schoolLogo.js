// School crest. The image file (school-logo.webp) sits next to this file.
// It starts as a full web address (so it also works inside the print window),
// then is swapped for an embedded copy once loaded so it keeps working offline.
import logoUrl from "./school-logo.webp";

const absolute = typeof window !== "undefined"
  ? new URL(logoUrl, window.location.origin).href
  : logoUrl;

export let SCHOOL_LOGO = absolute;

if (typeof window !== "undefined" && typeof fetch === "function") {
  fetch(logoUrl)
    .then(r => r.blob())
    .then(b => {
      const fr = new FileReader();
      fr.onload = () => { SCHOOL_LOGO = fr.result; };
      fr.readAsDataURL(b);
    })
    .catch(() => {});
}
