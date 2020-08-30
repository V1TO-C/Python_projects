let theme = localStorage.getItem("theme")
let contactButton = document.getElementById("contact-Btn");

if(theme == null){
    setTheme("light")
}else{
    setTheme(theme)
}

let themeDots = document.getElementsByClassName("theme-dot")


for (var i=0; themeDots.length > i; i++){
    themeDots[i].addEventListener("click", function(){
        let mode = this.dataset.mode
        console.log("Option clicked:", mode, i)
        setTheme(mode)
        let hrf = document.getElementById("theme-style").href
        console.log(hrf)
    })
}

function setTheme(mode){
    if(mode == "light"){
        document.getElementById("theme-style").href = "default.css"
    }

    if(mode == "blue"){
        document.getElementById("theme-style").href = "blue_theme.css"
    }

    if(mode == "green"){
        document.getElementById("theme-style").href = "green_theme.css"
    }

    if(mode == "purple"){
        document.getElementById("theme-style").href = "purple_theme.css"
    }

    localStorage.setItem("theme", mode)
}

contactButton.addEventListener("click", downFunction)

function downFunction(){
    window.scrollTo({
        top: document.body.scrollHeight,
        left: 0,
        behavior: 'smooth'
      });
}
