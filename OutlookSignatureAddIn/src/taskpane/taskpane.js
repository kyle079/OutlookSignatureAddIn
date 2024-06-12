const axios = require('axios');

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("saveTokenButton").addEventListener("click", saveDomainToken);
        document.getElementById("clearCacheButton").addEventListener("click", clearCache);
        document.getElementById("getSignaturesButton").addEventListener("click", getSignatures);
        initIndexedDB();
        refreshSignatures();
    }
});

function refreshSignatures() {
    setInterval(getSignatures, 1000 * 60 * 30); // Refresh every 30 minutes
}

function saveDomainToken() {
    const domainToken = document.getElementById("domainTokenInput").value;
    localStorage.setItem("domainToken", domainToken);
}

let db;

function initIndexedDB() {
    const request = indexedDB.open("SignatureDB", 1);

    request.onerror = (event) => console.error("IndexedDB error: ", event.target.errorCode);

    request.onsuccess = (event) => {
        db = event.target.result;
        console.log("IndexedDB initialized successfully.");
    };

    request.onupgradeneeded = (event) => {
        db = event.target.result;
        db.createObjectStore("signatures", { keyPath: "id", autoIncrement: true });
    };
}

function cacheSignatures(signatures) {
    if (!Array.isArray(signatures)) {
        console.error("Invalid signatures format:", signatures);
        return;
    }

    const transaction = db.transaction(["signatures"], "readwrite");
    const objectStore = transaction.objectStore("signatures");

    objectStore.clear().onsuccess = () => {
        signatures.forEach((signature, index) => {
            objectStore.add({ id: index, html: signature.html, name: signature.name });
        });
    };

    transaction.oncomplete = () => console.log("Signatures cached successfully.");
    transaction.onerror = (event) => console.error("IndexedDB transaction error: ", event.target.errorCode);
}

function clearCache() {
    const transaction = db.transaction(["signatures"], "readwrite");
    const objectStore = transaction.objectStore("signatures");

    objectStore.clear().onsuccess = () => {
        console.log("Cache cleared successfully.");
        document.getElementById("signaturesList").innerHTML = "";
        document.getElementById("resultMessage").textContent = "Cache cleared.";
    };

    transaction.onerror = (event) => console.error("IndexedDB transaction error: ", event.target.errorCode);
}

async function getSignatures() {
    document.getElementById("signaturesList").innerHTML = "";
    document.getElementById("resultMessage").textContent = "";
    document.getElementById("loading").style.display = "block";

    const signaturesHtmlArray = [];
    const transaction = db.transaction(["signatures"], "readonly");
    const objectStore = transaction.objectStore("signatures");

    objectStore.openCursor().onsuccess = async (event) => {
        const cursor = event.target.result;
        if (cursor) {
            signaturesHtmlArray.push({ html: cursor.value.html, name: cursor.value.name });
            cursor.continue();
        } else {
            if (signaturesHtmlArray.length) {
                setSignaturesHtml(signaturesHtmlArray);
                document.getElementById("resultMessage").textContent = "Retrieved from cache.";
            } else {
                await fetchSignaturesFromAPI();
                document.getElementById("resultMessage").textContent = "Retrieved from API.";
            }
            document.getElementById("loading").style.display = "none";
        }
    };

    transaction.onerror = async (event) => {
        console.error("IndexedDB transaction error: ", event.target.errorCode);
        await fetchSignaturesFromAPI();
    };
}

async function fetchSignaturesFromAPI() {
    const domainToken = localStorage.getItem("domainToken");
    if (!domainToken) {
        throw new Error("Domain token is not saved.");
    }
    const data = await fetchSignatures(domainToken);
    const processedHtml = await processSignatureHtml(data.signatures.signatures);
    cacheSignatures(processedHtml);
    setSignaturesHtml(processedHtml);
}

async function processSignatureHtml(signatures) {
    const signaturesHtmlArray = [];
    for (const signature of signatures) {
        if (!signature.html) {
            continue;
        }
        let signatureHtml = signature.html;

        if (signature.images && signature.images.length) {
            signature.images.forEach((image) => {
                const imgTagRegex = new RegExp(`<img([^>]*)src=["']?cid:${image.name}["']?([^>]*)>`, 'g');
                signatureHtml = signatureHtml.replace(imgTagRegex, `<img$1src="data:image/png;base64,${image.base64}"$2>`);
            });
        }

        signaturesHtmlArray.push({ name: signature.name, html: signatureHtml });
    }

    return signaturesHtmlArray;
}

async function setSignaturesHtml(signatures) {
    const signaturesList = document.getElementById("signaturesList");
    signaturesList.innerHTML = "";

    signatures.forEach((signature, index) => {
        const title = document.createElement("label");
        title.className = "ms-Label";
        title.textContent = signature.name;
        signaturesList.appendChild(title);

        const iframe = document.createElement("iframe");
        iframe.className = "signature-frame";
        iframe.id = `signatureFrame${index}`;
        signaturesList.appendChild(iframe);

        const iframeDoc = iframe.contentDocument || iframe.contentWindow.document;
        iframeDoc.open();
        iframeDoc.write(`<!DOCTYPE html><html><head><title>Signature</title></head><body>${signature.html}</body></html>`);
        iframeDoc.close();
    });
}

async function fetchSignatures(domainToken) {
    const { emailAddress } = Office.context.mailbox.userProfile;
    const { hostName, hostVersion, manifestVersion } = Office.context.mailbox.diagnostics;

    const config = {
        method: 'get',
        maxBodyLength: Infinity,
        url: `https://creatioleadapi.azure-api.net/sharepoint/ListDomainToken?domaintoken=${domainToken}&emailaddress=${emailAddress}&hostName=${hostName}&hostVersion=${hostVersion}&manifestVersion=${manifestVersion}`,
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
        }
    };

    try {
        const response = await axios.request(config);
        return response.data;
    } catch (error) {
        console.error(error);
        throw error;
    }
}
