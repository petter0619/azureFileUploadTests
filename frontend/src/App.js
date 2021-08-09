import LinkButton from "./LinkButton";
import AzureButton from "./AzureButton";

function App() {

  // Container level SAS key
  const fileUrl = "https://teststorageaccount0619.blob.core.windows.net/report-templates/ProjectRapport75fce9be-aa0e-49f7-8dcc-8153b94894e0.xlsx?sv=2020-08-04&se=2021-08-02T11%3A56%3A49Z&sr=b&sp=rd&sig=llRccsFlVvGb%2BA8K0Ts8DyVcvVB3ve9TH7wIgLu9p2k%3D";

  return (
    <div>
        <h2>Option 1: Programmatic Link Click</h2>
        <LinkButton downloadUrl={fileUrl}/>

        <h2>Option 2: Package</h2>
        <AzureButton downloadUrl={fileUrl}/>
    </div>
  );
}

export default App;
