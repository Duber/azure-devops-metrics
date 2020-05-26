import { o } from 'odata';
import { Headers } from 'node-fetch';
import { writeFile, utils } from 'xlsx'


(async () => {

    let config = {
        token: process.env.DEVOPS_API_TOKEN,
        org: process.env.DEVOPS_ORG
    }

    let url = `https://analytics.dev.azure.com/${config.org}/_odata/v3.0-preview`
    let httpConfig = {
        headers: new Headers({
            "Authorization": `Basic ${config.token}`
        }),
    }

    let projects = await o(url, httpConfig)
    .get('Projects')
    .query({ $select: "ProjectId" });

    let projectFilter = ""
    for(let i=0; i < projects.length; i++){
        if (i > 0)
            projectFilter = projectFilter + " or "
        projectFilter = projectFilter + `ProjectSK eq ${projects[i].ProjectId}`

    }

    let data = await o(url, httpConfig)
    .get('WorkItems')
    .query({ $select: "WorkItemId,Title,CreatedDateSK,InProgressDateSK,CompletedDateSK,WorkItemType"
                , $filter: `CompletedDateSK ne null and (${projectFilter})`
                , $expand: "Area($select=AreaName),Project($select=ProjectName)"});
    
    let sheetAoA = [["ID", "Link", "Name", "Backlog", "InProgress", "Done", "Type", "Project", "Area"]];
    console.log(`Found ${data.length} items`)
    data.forEach(function (item) {
        sheetAoA.push([item.WorkItemId, `https://dev.azure.com/${config.org}/${item.Project.ProjectName}/_workitems/edit/${item.WorkItemId}`, item.Title, item.CreatedDateSK, item.InProgressDateSK, item.CompletedDateSK, item.WorkItemType, item.Project.ProjectName, item.Area.AreaName])
    })

    var ws = utils.aoa_to_sheet(sheetAoA);
    let wb = utils.book_new();
    utils.book_append_sheet(wb, ws, "Sheet1");
    writeFile(wb, `${config.org}_${currentDate()}.xlsx`);

    console.log("END")
})();

function currentDate(){
    let today = new Date();
    return `${today.getFullYear()}${("0" + (today.getMonth() + 1)).slice(-2)}${("0" + today.getDate()).slice(-2)}`;
}