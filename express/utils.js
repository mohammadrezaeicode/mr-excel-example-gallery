export function hasImage(data) {
    const strData = JSON.stringify(data).toLowerCase()
    if(!strData) return false;
    return strData.indexOf("image") > -1;
}