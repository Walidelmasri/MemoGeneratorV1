namespace MemoGenerator.ViewModels
{
    public class MemoCreateVm
    {
        public string? To { get; set; }
        public string? Through { get; set; }   // optional
        public string? From { get; set; }
        public string? Subject { get; set; }
        public string? Body { get; set; }
        public string? Classification { get; set; }

        // Images: we’ll load defaults from wwwroot; no upload UI for now
        public bool UseDefaultImages { get; set; } = true;
    }
}
