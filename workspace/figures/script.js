var FIGURE_CLASS = "fortyPercent"

const detectRadio = input => {
    FIGURE_CLASS = input.value; //Set figure class for output from radio

    fig = generateFigure(document.getElementById('input').value); //update figure output with new figure class
    out.innerText = fig;
};

const detectTextArea = ta => {
    out = document.getElementById('output'); //Get textarea's value, trim for figure HTML output

    fig = generateFigure(ta.value);
    out.innerText = fig;
};

const trimTags = string => {
    if (string.includes('MadCap:conditions')) {
        let op = string.indexOf('>');
        let cl = string.slice(op,).indexOf('<') + op; //Get conditions out

        string = string.slice(op,cl);
    } else {
        string = string.split('<p>').join(''); //No conditions, just remove p tags
        string = string.split('</p>').join('');
    };

    return string;
};

const generateFigure = input => {
    let newline = input.indexOf('\n'); //Input should be two lines (graphic, caption) so separate by newline
    var conditions = '';  //Modify this with conditions if present

    const first = input.slice(0,newline);
    const second = input.slice(newline+1,).trim();

    if (first.includes('MadCap:conditions')) {
        let op = first.indexOf('MadCap:conditions');
        let cl = first.indexOf('>');

        conditions = first.slice(op,cl); //Get conditions out of tag, save for figure tag in output
    };

    //CODE RELIES ON GRAPHIC BEING FIRST LINE, CAPTION BEING SECOND
    const img = trimTags(first);
    const caption = trimTags(second);

    return `<figure class="${FIGURE_CLASS}" ${conditions}>\n<img src="../../Resources/Images/New/${img}.png"/>\n<figcaption>${caption}</figcaption>\n</figure>`
};

const copyToClipboard = () => {
    var copyText = document.getElementById('output');
    navigator.clipboard.writeText(copyText.innerText); //Copy textarea value
}