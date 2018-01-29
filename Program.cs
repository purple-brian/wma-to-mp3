using MediaFoundation;
using MediaFoundation.Misc;
using NAudio.Lame;
using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertWMAToMP3
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Please specify source directory and destination directory.");
                return;
            }

            // Start up media foundation.
            HResult hr;
            hr = MFExtern.MFStartup(0x10070, MFStartup.Full);
            MFError.ThrowExceptionForHR(hr);

            string sourceDirectory = args[0];
            string destDirectory = args[1];

            ProcessFiles(sourceDirectory, destDirectory);

            // Shut down media foundation.
            MFExtern.MFShutdown();
        }

        static void ProcessFiles(string sourceDirectory, string destDirectory)
        {
            if (!Directory.Exists(destDirectory))
            {
                Directory.CreateDirectory(destDirectory);
            }

            IEnumerable<string> wmaFiles = Directory.EnumerateFiles(sourceDirectory, "*.wma");
            foreach (string wmaFile in wmaFiles)
            {
                string mp3File = Path.Combine(destDirectory, Path.ChangeExtension(Path.GetFileName(wmaFile), "mp3"));
                Console.Write($"Converting '{wmaFile}'... ");
                ConvertFile(wmaFile, mp3File);
                Console.WriteLine($"wrote '{mp3File}'.");
            }

            IEnumerable<string> directories = Directory.EnumerateDirectories(sourceDirectory);
            foreach (string directory in directories)
            {
                string directoryName = Path.GetFileName(directory);
                ProcessFiles(directory, Path.Combine(destDirectory, directoryName));
            }
        }

        static void ConvertFile(string sourceFileName, string destFileName)
        {
            IMFMediaSource mediaSource = GetMediaSource(sourceFileName);
            int bitRate = GetBitRate(mediaSource);
            IMFMetadata metadata = GetMetadata(mediaSource);

            ID3TagData tagData = new ID3TagData()
            {
                Title = GetStringProperty(metadata, "Title"),
                Artist = GetStringProperty(metadata, "Author"),
                Album = GetStringProperty(metadata, "WM/AlbumTitle"),
                Year = GetStringProperty(metadata, "WM/Year"),
                Genre = GetStringProperty(metadata, "WM/Genre"),
                Track = GetUIntProperty(metadata, "WM/TrackNumber"),
                AlbumArtist = GetStringProperty(metadata, "WM/AlbumArtist")
            };
            COMBase.SafeRelease(metadata);
            metadata = null;
            COMBase.SafeRelease(mediaSource);
            mediaSource = null;

            using (AudioFileReader reader = new AudioFileReader(sourceFileName))
            {
                using (LameMP3FileWriter writer = new LameMP3FileWriter(destFileName, reader.WaveFormat, bitRate, tagData))
                {
                    reader.CopyTo(writer);
                }
            }
        }

        static IMFMediaSource GetMediaSource(string sourceFileName)
        {
            HResult hr;

            // Get an IMFMediaSource.
            IMFSourceResolver sourceResolver;
            hr = MFExtern.MFCreateSourceResolver(out sourceResolver);
            MFError.ThrowExceptionForHR(hr);
            MFObjectType objectType = MFObjectType.Invalid;
            object source;
            hr = sourceResolver.CreateObjectFromURL(sourceFileName, MFResolution.MediaSource, null, out objectType, out source);
            MFError.ThrowExceptionForHR(hr);
            IMFMediaSource mediaSource = (IMFMediaSource)source;
            COMBase.SafeRelease(sourceResolver);
            sourceResolver = null;

            return mediaSource;
        }

        static IMFMetadata GetMetadata(IMFMediaSource mediaSource)
        {
            HResult hr;

            // Get IMFPresentationDescriptor.
            IMFPresentationDescriptor presentationDescriptor;
            hr = mediaSource.CreatePresentationDescriptor(out presentationDescriptor);
            MFError.ThrowExceptionForHR(hr);

            // Get IMFMetadataProvider.
            object provider;
            hr = MFExtern.MFGetService(mediaSource, MFServices.MF_METADATA_PROVIDER_SERVICE, typeof(IMFMetadataProvider).GUID, out provider);
            MFError.ThrowExceptionForHR(hr);
            IMFMetadataProvider metadataProvider = (IMFMetadataProvider)provider;

            // Get IMFMetadata.
            IMFMetadata metadata;
            hr = metadataProvider.GetMFMetadata(presentationDescriptor, 0, 0, out metadata);
            MFError.ThrowExceptionForHR(hr);
            COMBase.SafeRelease(presentationDescriptor);
            presentationDescriptor = null;
            COMBase.SafeRelease(metadataProvider);
            metadataProvider = null;

            return metadata;
        }

        static int GetBitRate(IMFMediaSource mediaSource)
        {
            HResult hr;

            // Get IMFPresentationDescriptor.
            IMFPresentationDescriptor presentationDescriptor;
            hr = mediaSource.CreatePresentationDescriptor(out presentationDescriptor);
            MFError.ThrowExceptionForHR(hr);

            // Get IMFStreamDescriptor.
            bool isStreamSelected;
            IMFStreamDescriptor streamDescriptor;
            hr = presentationDescriptor.GetStreamDescriptorByIndex(0, out isStreamSelected, out streamDescriptor);
            MFError.ThrowExceptionForHR(hr);

            // Get bit rate.
            int bitRate;
            hr = streamDescriptor.GetUINT32(MFAttributesClsid.MF_SD_ASF_EXTSTRMPROP_AVG_DATA_BITRATE, out bitRate);
            MFError.ThrowExceptionForHR(hr);
            bitRate /= 1000;
            if (bitRate <= 0)
            {
                throw new ApplicationException($"Invalid bit rate: {bitRate}");
            }

            return bitRate;
        }

        static string GetStringProperty(IMFMetadata metadata, string name)
        {
            PropVariant value = new PropVariant();
            metadata.GetProperty(name, value);
            return value.GetVariantType() == ConstPropVariant.VariantType.String ? value.GetString() : null;
        }

        static string GetUIntProperty(IMFMetadata metadata, string name)
        {
            PropVariant value = new PropVariant();
            metadata.GetProperty(name, value);
            return value.GetVariantType() == ConstPropVariant.VariantType.UInt32 ? value.GetUInt().ToString() : null;
        }
    }
}
