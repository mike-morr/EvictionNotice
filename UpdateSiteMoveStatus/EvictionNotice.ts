
let countDown = 8 // Controls number of seconds until automatic redirect

const hFunc = () => document.createElement("h2")
const pFunc = () => document.createElement("p")
const header = document.createElement("h1")
const h2 = hFunc()
const h3 = hFunc()
let defaultUrl = "http://lab.yosp.io/sites/search/Pages/results.aspx?k="
let newUrl = null

function getRedirectFromSharePoint() {
	hideContent() // Hide everything on the page
	
	// Load the jQuery
	// TODO: See if there is a better way to reuse the jQuery that is already loaded with SharePoint
	// This would reduce the initial page flicker
	let script = document.createElement("script");
	script.src = 'http://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.2.4.min.js';
	script.type = 'text/javascript';
	
	// When script loads kick off the SP REST Query and either fail or process response
	// TODO: Change the REST query so that it brings back all of the links that "StartWith" the old url
	// This would allow us to sort and use the longest URL that matches, a.k.a. most specific match
	script.onload = function () {
		let settings: JQueryAjaxSettings = {
			headers: { "Accept": "application/json;odata=verbose" },
			method: "GET",
			url: `http://lab.yosp.io/sites/dev/_api/web/lists/getbytitle('SiteMoves')/items?$filter=OldUrl eq '${encodeURIComponent(window.location.href)}'`,
			error: (r) => processFailure(r),
			success: (r) => {
				processResponse(r.d.results, false)
			}
		}
		// Actually calls the server, the next line of code to run will
		// be the success or error handler above.
		$.ajax(settings)

	}
	document.getElementsByTagName("head")[0].appendChild(script)

}

function hideContent() {
	document.getElementById("s4-workspace").setAttribute("style", "display: none;")
}

function processFailure(response) {
	let code = response.status
	let message = response.statusText
	console.log(`Error in http request:\r\nError Code: ${code}\r\nError Message: ${message}`);
	processResponse(response, true)
}

function processResponse(response, error: boolean) {
	header.innerHTML = "Eviction Notice!";
	
	if (error || !response || !response.NewUrl) {
		// Handle URL not found
		defaultUrl = defaultUrl + window.location.pathname.split("/").join(" ").trim()
		h2.innerHTML = `The previous tenant did not leave a forwarding address!  Let's see if we can find it by using search at ${defaultUrl}`
		h3.innerHTML = `You will be redirected in ${countDown}`
	} else {
		// Handle URL was found
		newUrl = response.NewUrl
		h2.innerHTML = `This previous tenant left a forwarding address!  This site has been permanently moved to ${newUrl}`
		h3.innerHTML = `You will be redirected in ${countDown}`
	}

	// Set styles
	header.setAttribute("style", "margin: 30px 30px;")
	h2.setAttribute("style", "margin: 30px 30px;")
	h3.setAttribute("style", "margin: 30px 30px;")

	// Add our elements to the end of the body
	document.body.appendChild(header)
	document.body.appendChild(h2)
	document.body.appendChild(pFunc())
	document.body.appendChild(h3)

	let redirectCounter = setInterval(() => {
		if (countDown === 0) {
			clearInterval(redirectCounter)
			window.location.href = newUrl || defaultUrl
		}
		h3.innerHTML = `You will be redirected in ${countDown}`
		countDown--
	}, 1000)

}

//let handle = setInterval(() => {
//	if (document.getElementById("s4-workspace") !== null) {
//		startup()
//	}
//}, 60)

// Here we wait until sp.js is loaded, then we start talking to SharePoint
// Note TypeScript is just syntax sugar on top of JavaScript.  So the following line
// will not be in the compilation, it is only to make the compiler happy.  So that ...
declare var ExecuteOrDelayUntilScriptLoaded

// ...this line does not throw any type errors.
// TODO: Tweak to hide content faster, I really don't like the initial page flicker, it feels jarring
ExecuteOrDelayUntilScriptLoaded(getRedirectFromSharePoint, "cui.js");