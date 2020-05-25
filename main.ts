import { o } from 'odata';
import { Headers } from 'node-fetch';
import { writeFile, utils } from 'xlsx'


(async () => {
    let authToken = "<TOKEN>"
    let org = "<ORG_NAME>"
    let url = `https://analytics.dev.azure.com/${org}/_odata/v3.0-preview`
    let config = {
        headers: new Headers({
            "Authorization": `Basic ${authToken}`
        }),
    }

    let projects = await o(url, config)
    .get('Projects')
    .query({ $select: "ProjectId" });

    let projectFilter = ""
    for(let i=0; i < projects.length; i++){
        if (i > 0)
            projectFilter = projectFilter + " or "
        projectFilter = projectFilter + `ProjectSK eq ${projects[i].ProjectId}`

    }

    let data = await o(url, config)
    .get('WorkItems')
    .query({ $select: "WorkItemId,Title,CreatedDateSK,InProgressDateSK,CompletedDateSK,WorkItemType"
                , $filter: `CompletedDateSK ne null and (${projectFilter})`
                , $expand: "Area($select=AreaName),Project($select=ProjectName)"});
    
    let sheetAoA = [["ID", "Link", "Name", "Backlog", "InProgress", "Done", "Type", "Project", "Area"]];
    console.log(`Found ${data.length} items`)
    data.forEach(function (item) {
        sheetAoA.push([item.WorkItemId, `https://dev.azure.com/${org}/${item.Project.ProjectName}/_workitems/edit/${item.WorkItemId}`, item.Title, item.CreatedDateSK, item.InProgressDateSK, item.CompletedDateSK, item.WorkItemType, item.Project.ProjectName, item.Area.AreaName])
    })

    var ws = utils.aoa_to_sheet(sheetAoA);
    let wb = utils.book_new();
    utils.book_append_sheet(wb, ws, "Sheet1");
    writeFile(wb, "out.xlsx");

    console.log("END")
})();