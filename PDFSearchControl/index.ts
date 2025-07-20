import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class AccountSearchControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _context: ComponentFramework.Context<IInputs>;
    private _accountId: string | undefined;


    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._context = context;
        this._container = container;
        this._accountId = (context.mode as any).contextInfo.entityId;
        this._container.innerHTML = "<p>Loading PDF...</p>";
        this.loadPdf();
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        this._accountId = (context.mode as any).contextInfo.entityId;
        this._container.innerHTML = "<p>Refreshing PDF...</p>";
        this.loadPdf();
    }

    private async loadPdf(): Promise<void> {
        if (!this._accountId) {
            this._container.innerHTML = "<p>No Account ID available.</p>";
            return;
        }

        try {
            const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();
            const fileUrl = `${clientUrl}/api/data/v9.0/accounts(${this._accountId})/ram_pdfattachment/$value`;

            const response = await fetch(fileUrl, {
                method: "GET",
                headers: { "Accept": "application/pdf" },
                credentials: "include"
            });

            if (!response.ok) {
                this._container.innerHTML = "<p>No File Found.</p>";
                return;
            }

            const blob = await response.blob();
            const base64Pdf = await this.blobToBase64(blob);

            const pdfSrc = `data:application/pdf;base64,${base64Pdf}`;
            this._container.innerHTML = `
                <iframe src="${pdfSrc}" width="100%" height="600px"
                        style="border:1px solid #ccc; border-radius:4px;" />
            `;
        } catch (error: unknown) {
            if (error instanceof Error) {
                this._container.innerHTML = `<p>Error loading PDF: ${error.message}</p>`;
            } else {
                this._container.innerHTML = `<p>An unexpected error occurred.</p>`;
            }
        }
    }

    private blobToBase64(blob: Blob): Promise<string> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => {
                const dataUrl = reader.result as string;
                const base64 = dataUrl.split(",")[1];
                resolve(base64);
            };
            reader.onerror = reject;
            reader.readAsDataURL(blob);
        });
    }


    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}
