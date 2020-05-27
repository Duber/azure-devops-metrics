import { writeFile, utils } from 'xlsx'
import axios from 'axios'

(async () => {

    let config = {
        token: process.env.DEVOPS_API_TOKEN,
        org: process.env.DEVOPS_ORG
    }

    let url = `https://analytics.dev.azure.com/${config.org}/_odata/v3.0-preview`

    let axiosOptions = {
        auth: {
            username: "user",
            password: config.token
        }
    }

    let numItems = 0;

    let projectFilter = await buildProjectsFilter(url, axiosOptions);
    
    let itemsQuery = `${url}/WorkItems?
    $filter=CompletedDateSK ne null and (${projectFilter})
    &$select=WorkItemId,Title,CreatedDateSK,InProgressDateSK,CompletedDateSK,WorkItemType
    &$expand=Area($select=AreaName),Project($select=ProjectName)`


    let itemsResponse = await axios.get(itemsQuery, axiosOptions)
    let items = itemsResponse.data.value
    
    let sheetAoA = [["ID", "Link", "Name", "Backlog", "InProgress", "Done", "Type", "Project", "Area"]];
    addItemsToSheet(items, sheetAoA, config);
    numItems += items.length

    while(itemsResponse.data.hasOwnProperty("@odata.nextLink")){
        itemsResponse = await axios.get(itemsResponse.data["@odata.nextLink"], axiosOptions)
        items = itemsResponse.data.value
        addItemsToSheet(items, sheetAoA, config)
        numItems += items.length
    }
    console.log(`Found ${numItems} items`)

    writeExcel(sheetAoA, config);

    console.log("END")
})();

function addItemsToSheet(items: any, sheetAoA: string[][], config: { token: string; org: string; }) {
    items.forEach(function (item) {
        sheetAoA.push([item.WorkItemId, `https://dev.azure.com/${config.org}/${item.Project.ProjectName}/_workitems/edit/${item.WorkItemId}`, item.Title, item.CreatedDateSK, item.InProgressDateSK, item.CompletedDateSK, item.WorkItemType, item.Project.ProjectName, item.Area.AreaName]);
    });
}

async function buildProjectsFilter(url: string, axiosOptions: { auth: { username: string; password: string; }; }) {
    let projectsQuery = `${url}/Projects?$select=ProjectId`;
    let projectsResponse = await axios.get(projectsQuery, axiosOptions);
    let projects = projectsResponse.data.value;
    let projectFilter = "";
    for (let i = 0; i < projects.length; i++) {
        if (i > 0)
            projectFilter = projectFilter + " or ";
        projectFilter = projectFilter + `ProjectSK eq ${projects[i].ProjectId}`;
    }
    return projectFilter;
}

function writeExcel(sheetAoA: string[][], config: { token: string; org: string; }) {
    var ws = utils.aoa_to_sheet(sheetAoA);
    let wb = utils.book_new();
    utils.book_append_sheet(wb, ws, "Sheet1");
    writeFile(wb, `${config.org}_${currentDate()}.xlsx`);
}

function currentDate(){
    let today = new Date();
    return `${today.getFullYear()}${("0" + (today.getMonth() + 1)).slice(-2)}${("0" + today.getDate()).slice(-2)}`;
}