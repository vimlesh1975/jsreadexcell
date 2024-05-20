var filter1 = document.getElementsByTagName('defs')[0];
filter1.innerHTML = '<filter id="filter1" x="0" y="0" width="100%" height="100%"> <feSpecularLighting result="spec1" specularExponent="12" lighting-color="yellow"> <fePointLight x="0" y="0" z="14" > <animate attributeName="x" values="-467.5;517.5;517.5" keyTimes="0;0.5; 1" dur="3s" repeatCount="indefinite" /> </fePointLight> </feSpecularLighting> <feComposite in="SourceGraphic" in2="spec1" operator="arithmetic" k1="0" k2="1" k3="1" k4="0" /> </filter>'
var allRectangles = document.getElementsByTagName('rect');

Array.from(allRectangles).forEach((element) => {
    element.style.filter = 'url(#filter1)';
});

const scriptgsap = document.createElement("script");
scriptgsap.src = "./js/gsap.min.js";
scriptgsap.setAttribute("type", "text/javascript");

var elements = document.querySelectorAll('rect, image, text, path, circle');

const sortedElements = Array.from(elements).sort(function (a, b) {
    return a.getBoundingClientRect().top - b.getBoundingClientRect().top;
});

const inAnimation=()=>{
 
    document.body.style.opacity = 1;
    sortedElements.forEach((element, index) => {
        var pathTransform = 0;
        if (element.tagName === 'path') {
            pathTransform = element.transform.animVal[0].matrix.e
        }
        const scalefactor = element.parentNode.getCTM().a;
        gsap.set(element, { x: -2100 / scalefactor, opacity: 0 });
        gsap.to(element, {
            x: (element.tagName === 'path') ? pathTransform : 0,
            opacity: 1,
            duration: 0.5,
            delay: index * 0.03,
            ease: "",
        });
    });

   
}

const outAnimation = () => {
    sortedElements.forEach((element, index) => {
        const scalefactor = element.parentNode.getCTM().a;
        gsap.set(element, { x: -2100 / scalefactor, opacity: 0 });
        gsap.from(element, {
            x: 0,
            opacity: 1,
            duration: 0.5,
            delay: (sortedElements.length - index - 1) * 0.03,
            ease: "power2.out"
        });
    });

}

scriptgsap.onload = function () {
    setTimeout(() => {
        // timeout is nessaesaary to set all variable set by client.
        inAnimation();
    }, 100);
 
};



document.body.appendChild(scriptgsap);



const excelRead = document.createElement("script");
excelRead.src = "./xlsx.full.min.js";
excelRead.setAttribute("type", "text/javascript");

excelRead.onload = function () {
    // Function to read Excel file from the same directory
    function readExcelFromSameDirectory(filename, callback) {
        const xhr = new XMLHttpRequest();
        xhr.open('GET', filename, true);
        xhr.responseType = 'arraybuffer';

        xhr.onload = function (e) {
            if (xhr.status === 200) {
                const data = new Uint8Array(xhr.response);
                const workbook = XLSX.read(data, { type: 'array' });

                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet);

                callback(null, jsonData);
            } else {
                callback(new Error('Failed to load Excel file'));
            }
        };

        xhr.send();
    }

    // Usage
    readExcelFromSameDirectory('./aa.xlsx', (err, excelData) => {
        if (err) {
            console.error('Error reading Excel file:', err);
            return;
        }
        console.log('Excel data:', excelData);
        var i=0;
        setInterval(() => {
            setTimeout(() => {
                updatestring('name', excelData[i].Name);
                updatestring('age', excelData[i].Age);
                inAnimation()
            }, 1000);
        
            i++
            if (i>excelData.length-1){
                i=0;
            }
        }, 3000);

    });
};

document.body.appendChild(excelRead);


