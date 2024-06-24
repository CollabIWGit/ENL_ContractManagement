// OwnerDashboard.ts
export default class DespatcherDashboardObj {
    private htmlContent: string;

    constructor() {
        this.htmlContent = `
            <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.3.1/dist/css/bootstrap.min.css" integrity="sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T" crossorigin="anonymous">
            <link href="https://nightly.datatables.net/css/jquery.dataTables.css" rel="stylesheet" type="text/css" />
            <script src="https://nightly.datatables.net/js/jquery.dataTables.js"></script>

            <style>

                /* Styling for datatable search bar */

                div.dt-container .dt-search {
                    padding-right: 1rem;
                }

                div.dt-container .dt-search input {
                    border: none;
                    background: none;
                    border-bottom: 0.1rem solid #062470;
                    margin-bottom: 1rem;
                    width: 40%;
                }

                div.dt-container .dt-search input::placeholder {
                    color: #aaa;
                }
                
                div.dt-container .dt-search label {
                    display: none;
                }

                div.dt-container .dt-search input:focus {
                    outline: none;
                    box-shadow: none;
                }

                /* Styling for datatable dropdown */

                div.dt-container select.dt-input {
                    border: none;
                    border-bottom: 0.1rem solid #062470;
                    margin-right: 0.5rem;
                }

                div.dt-container select.dt-input:focus {
                    outline: none;
                }

                /* Styling for datatable pagination */

                div.dt-container .dt-paging .dt-paging-button {
                    border-radius: 50%;
                    background-color: transparent;
                }

                div.dt-container .dt-paging .dt-paging-button:hover {
                    border-radius: 50%;
                    background: rgb(239, 125, 23, 0.9);
                    border: none;
                }

                div.dt-container .dt-paging .dt-paging-button.disabled {
                    background-color: transparent;
                    background: transparent;
                }

                div.dt-container .dt-paging .dt-paging-button.current, div.dt-container .dt-paging .dt-paging-button.current:hover {
                    border-radius: 50%;
                    background-color: #ef7d17;
                    border: none;
                }

                /*DataTable css*/
                div.dt-container.dt-empty-footer tbody > tr:last-child > * {
                    border-bottom: 1px solid #062470;
                }
                
                table.dataTable > thead > tr > th, table.dataTable > thead > tr > td {
                    border-bottom: 1px solid #062470;
                }


            </style>
            
            <div class="wrapper d-flex align-items-stretch">

            <div id="contractsDatatableDiv"></div>
                
            </div>

        `;
    }

    public getHtmlContent(): string {
        return this.htmlContent;
    }
}
