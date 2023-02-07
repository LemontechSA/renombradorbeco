namespace renombradorbeco.app
{
    public class ContainerInfo<T> where T : FileSystemInfo
    {

        public readonly DirectoryInfo Value;

        public ContainerInfo(DirectoryInfo input)
        {
            this.Value = input;
        }

        public static implicit operator ContainerInfo<T>(string path)
        {
            var dirInfo = new DirectoryInfo(path);
            return new ContainerInfo<T>(dirInfo);
        }
    }
}