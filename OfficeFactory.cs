using System.Threading.Tasks;

namespace Office.Service
{
    public class OfficeFactory
    {
        private readonly OfficeSettings settings;
        public OfficeFactory(OfficeSettings officeSettings)
        {
            this.settings = officeSettings;
        }

        public async Task<BaseService> BuildAsync()
        {
            switch (settings.Type)
            {
                case TypeEnum.Manual:
                    return new ManualService(settings);
                case TypeEnum.Spir:
                    return new SpireService(settings);
                case TypeEnum.Aspose:
                    return new AsposeService(settings);
            }
            return await Task.FromResult(default(BaseService));
        }
    }
}
