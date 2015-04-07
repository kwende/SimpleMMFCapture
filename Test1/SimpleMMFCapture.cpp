// Test1.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
/* Include files to use the OpenCV API. */
#include <opencv/cv.h>
#include <opencv/cxcore.h>
#include <opencv/highgui.h>

/* Include files to use the MFPLAYER API. */
#include <mfapi.h>
#include <mfidl.h>
#include <Mfreadwrite.h>
#include <mferror.h>

/* Include files to use IFileOpenDialog. */
#include <shobjidl.h>   
#include <shlwapi.h>

#pragma comment(lib, "mfreadwrite")
#pragma comment(lib, "mfplat")
#pragma comment(lib, "mfuuid")

#define GRAY_CHANNELS	 ( 1 )
#define RGB24_CHANNELS	 ( 3 )
#define RGB32_CHANNELS	 ( 4 )

#define IF_FAIL_GOTO(hr) if (FAILED(hr)) { goto done; }

class WriteAVI
{
    /***********************************************************************/
    /***********************************************************************/
private:
    /*********** Variable **********/
    UINT64		rtStart;
    UINT64		rtDuration;
    DWORD		stream;

    /************ Object ***********/
    IMFSinkWriter	*pSinkWriter;
    GUID			VideoInputFormat;

    /*********** Function **********/
    template <class T> void SafeRelease(T **ppT);
    HRESULT		InitializeSinkWriter(IMFSinkWriter **ppSinkWriter,
        DWORD *pStream,
        LPCWSTR FileName,
        GUID FormatVideo,
        UINT32 BitRate,
        UINT32 Width,
        UINT32 Height,
        UINT32 Fps,
        INT Channels);
    HRESULT		WriteFrame(BYTE *pData,
        const LONGLONG& artStart,
        const LONGLONG& artDuration);

    /***********************************************************************/
    /***********************************************************************/
public:
    /*********** Variable **********/
    LONG		widthStep;
    DWORD		imageSize;
    INT			channels;
    UINT32		fps;
    UINT32		bitRate;
    UINT32		videoWidth;
    UINT32		videoHeight;

    /************ Object ***********/

    /*********** Function **********/
    void		WriteImage(BYTE *pData);
    bool		CreateVideo(LPCWSTR FileName,					/* Video file names.		*/
        GUID FormatVideo,					/* Video format.			*/
        UINT32 BitRate,					/* Bit rate.				*/
        UINT32 Width,						/* Width of video.			*/
        UINT32 Height,					/* Height of video.			*/
        UINT32 Fps,						/* Frame per second.		*/
        INT Channels);					/* Channels of video.		*/
    void		ReleaseVideo(void);

    WriteAVI(void)
    {
        /* Set default */
        pSinkWriter = NULL;
    }
    ~WriteAVI(void)
    {
    }
};

bool WriteAVI::CreateVideo(LPCWSTR FileName,
    GUID FormatVideo,
    UINT32 BitRate,
    UINT32 Width,
    UINT32 Height,
    UINT32 Fps,
    INT Channels)
{
    //HRESULT hr;
    HRESULT hr = CoInitializeEx(0, COINIT_APARTMENTTHREADED);

    /* Initialize the Media Foundation platform. */
    IF_FAIL_GOTO(hr = MFStartup(MF_VERSION));

    /* Initialize sink writer */
    IF_FAIL_GOTO(hr = InitializeSinkWriter(&pSinkWriter,
        &stream,
        FileName,
        FormatVideo,
        BitRate,
        Width,
        Height,
        Fps,
        Channels));
    if (SUCCEEDED(hr))
    {
        /* Set start point of the video frame. */
        rtStart = 0;

        /* Get the video attribute. */
        fps = Fps;
        bitRate = BitRate;
        channels = Channels;
        videoWidth = Width;
        videoHeight = Height;
        widthStep = Width * Channels;
        imageSize = Height * widthStep;

        /* Converts a video frame rate into a frame duration. */
        MFFrameRateToAverageTimePerFrame(Fps, 1, &rtDuration);

        /* Return true for success connecting. */
        return TRUE;
    }
    else
    {
        /* Shut down Media Foundation. */
        MFShutdown();
    }
done:
    /* Return false for fail connecting. */
    return FALSE;
}

HRESULT WriteAVI::InitializeSinkWriter(IMFSinkWriter **ppSinkWriter,
    DWORD *pStream,
    LPCWSTR FileName,
    GUID FormatVideo,
    UINT32 BitRate,
    UINT32 Width,
    UINT32 Height,
    UINT32 Fps,
    INT Channels)
{
    /* Initial variable */
    HRESULT hr;
    IMFSinkWriter   *pbufSinkWriter = NULL;
    IMFMediaType    *pMediaTypeOut = NULL;
    IMFMediaType    *pMediaTypeIn = NULL;
    DWORD           streamIndex;

    *ppSinkWriter = NULL;
    *pStream = NULL;

    /* Creates the sink writer from a URL or byte stream. */
    IF_FAIL_GOTO(hr = MFCreateSinkWriterFromURL(FileName, NULL, NULL, &pbufSinkWriter));

    /*-------------- SET OUTPUT MEDIA TYPE. --------------*/

    /* Creates an empty media type. */
    IF_FAIL_GOTO(hr = MFCreateMediaType(&pMediaTypeOut));

    /* Set the major type is video type. */
    IF_FAIL_GOTO(hr = pMediaTypeOut->SetGUID(MF_MT_MAJOR_TYPE, MFMediaType_Video));

    /* Set the encoding output format. */
    IF_FAIL_GOTO(hr = pMediaTypeOut->SetGUID(MF_MT_SUBTYPE, FormatVideo));

    /* Set the video bit rate. */
    IF_FAIL_GOTO(hr = pMediaTypeOut->SetUINT32(MF_MT_AVG_BITRATE, BitRate));

    /* Set the video interlace progressive. */
    IF_FAIL_GOTO(hr = pMediaTypeOut->SetUINT32(MF_MT_INTERLACE_MODE, MFVideoInterlace_Progressive));

    /* Set the frame size. */
    IF_FAIL_GOTO(hr = MFSetAttributeSize(pMediaTypeOut, MF_MT_FRAME_SIZE, Width, Height));

    /* Set the frame per second. */
    IF_FAIL_GOTO(hr = MFSetAttributeRatio(pMediaTypeOut, MF_MT_FRAME_RATE, Fps, 1));

    /* Set the pixel aspect ratio. */
    IF_FAIL_GOTO(hr = MFSetAttributeRatio(pMediaTypeOut, MF_MT_PIXEL_ASPECT_RATIO, 1, 1));

    /* Adds a stream to the sink writer. */
    IF_FAIL_GOTO(hr = pbufSinkWriter->AddStream(pMediaTypeOut, &streamIndex));


    /*--------------- SET INPUT MEDIA TYPE. ---------------*/

    /* Creates an empty media type. */
    IF_FAIL_GOTO(hr = MFCreateMediaType(&pMediaTypeIn));

    /* Set the major type is video type. */
    IF_FAIL_GOTO(hr = pMediaTypeIn->SetGUID(MF_MT_MAJOR_TYPE, MFMediaType_Video));

    /* Set the encoding input format. */
    if (SUCCEEDED(hr))
    {
        switch (Channels)
        {
        case GRAY_CHANNELS:	VideoInputFormat = MFVideoFormat_RGB8;   //******//
            break;
        case RGB24_CHANNELS:	VideoInputFormat = MFVideoFormat_RGB24;
            break;
        case RGB32_CHANNELS:	VideoInputFormat = MFVideoFormat_RGB32;
            break;
        default:				printf("Error InitializeSinkWriter : Cannot find specific channels.\n");
            goto done;
        }
        hr = pMediaTypeIn->SetGUID(MF_MT_SUBTYPE, VideoInputFormat);
    }
    else
    {
        goto done;
    }

    /* Set the video interlace progressive. */
    IF_FAIL_GOTO(hr = pMediaTypeIn->SetUINT32(MF_MT_INTERLACE_MODE, MFVideoInterlace_Progressive));

    /* Set the frame size. */
    IF_FAIL_GOTO(hr = MFSetAttributeSize(pMediaTypeIn, MF_MT_FRAME_SIZE, Width, Height));

    /* Set the frame per second. */
    IF_FAIL_GOTO(hr = MFSetAttributeRatio(pMediaTypeIn, MF_MT_FRAME_RATE, Fps, 1));

    /* Set the pixel aspect ratio. */
    IF_FAIL_GOTO(hr = MFSetAttributeRatio(pMediaTypeIn, MF_MT_PIXEL_ASPECT_RATIO, 1, 1));

    /* Sets the input format for a stream on the sink writer. */
    IF_FAIL_GOTO(hr = pbufSinkWriter->SetInputMediaType(streamIndex, pMediaTypeIn, NULL));


    /* Tell the sink writer to start accepting data. */
    IF_FAIL_GOTO(hr = pbufSinkWriter->BeginWriting());

    /* Return the pointer to the caller. */
    if (SUCCEEDED(hr))
    {
        *ppSinkWriter = pbufSinkWriter;
        (*ppSinkWriter)->AddRef();
        *pStream = streamIndex;
    }

done:

    /* Deallocate memory */
    SafeRelease(&pbufSinkWriter);
    SafeRelease(&pMediaTypeOut);
    SafeRelease(&pMediaTypeIn);
    return hr;
}

void WriteAVI::WriteImage(BYTE *pData)
{
    HRESULT hr = WriteFrame(pData, rtStart, rtDuration);
    if (SUCCEEDED(hr))
    {
        rtStart += rtDuration;
    }
}

HRESULT WriteAVI::WriteFrame(BYTE *pData, const LONGLONG& artStart, const LONGLONG& artDuration)
{
    /* Initial variable */
    HRESULT			hr;
    BYTE			*pBufData = NULL;
    IMFSample		*pSample = NULL;
    IMFMediaBuffer  *pBuffer = NULL;

    /* Create a new memory buffer. */
    IF_FAIL_GOTO(hr = MFCreateMemoryBuffer(imageSize, &pBuffer));

    /* Lock the buffer. */
    IF_FAIL_GOTO(hr = pBuffer->Lock(&pBufData, NULL, NULL));

    /* copy the video frame to the buffer. */
    IF_FAIL_GOTO(hr = MFCopyImage(pBufData,						/* Destination buffer.			*/
        widthStep,						/* Destination stride.			*/
        pData,							/* First row in source image.	*/
        widthStep,						/* Source stride.				*/
        widthStep,                      /* Image width in bytes.		*/
        videoHeight));				/* Image height in pixels.		*/

    /* Unlock the buffer. */
    if (pBuffer)
    {
        pBuffer->Unlock();
    }

    /* Set the data length of the buffer. */
    IF_FAIL_GOTO(hr = pBuffer->SetCurrentLength(imageSize));

    /* Create a media sample */
    IF_FAIL_GOTO(hr = MFCreateSample(&pSample));

    /* Add the buffer to the sample. */
    IF_FAIL_GOTO(hr = pSample->AddBuffer(pBuffer));

    /* Set the time stamp. */
    IF_FAIL_GOTO(hr = pSample->SetSampleTime(artStart));

    /* Set the time duration. */
    IF_FAIL_GOTO(hr = pSample->SetSampleDuration(artDuration));

    /* Send the sample to the Sink Writer. */
    IF_FAIL_GOTO(hr = pSinkWriter->WriteSample(stream, pSample));

done:

    /* Deallocate memory */
    SafeRelease(&pSample);
    SafeRelease(&pBuffer);
    return hr;
}

void WriteAVI::ReleaseVideo()
{
    /* End the writer video */
    pSinkWriter->Finalize();

    /* Deallocate memory */
    SafeRelease(&pSinkWriter);

    /* Shut down Media Foundation. */
    MFShutdown();

    CoUninitialize();
}

template <class T> void WriteAVI::SafeRelease(T **ppT)
{
    if (*ppT)
    {
        (*ppT)->Release();
        *ppT = NULL;
    }
}


//---------------------MAIN-------------------//
void main(void)
{
    WriteAVI c_WriteAVI;

    /* Initializes the COM library for use by the calling thread. */
    //HRESULT hr = CoInitializeEx(0, COINIT_APARTMENTTHREADED);
    HRESULT hr = S_OK;
    if (SUCCEEDED(hr))
    {
        /* Print the success text. */
        printf("Initial COM runtime success\n");

        /* Connect camera */
        IplImage *img;
        CvCapture *cap = cvCaptureFromCAM(0);
        IplImage *g_img = cvCreateImage(cvSize(640, 480), 8, 1);

        /* Convert char to LPCWSTR */
        size_t   size;
        char     FileName[] = "output.wmv";
        wchar_t	 *lpFileName = (wchar_t *)malloc(30 * sizeof(wchar_t));
        size = mbstowcs(lpFileName, FileName, 30);

        /* Create video */
        c_WriteAVI.CreateVideo(lpFileName,
            MFVideoFormat_WMV3,
            800000,
            640,
            480,
            30,
            3);

        int time = GetTickCount();
        while (TRUE)
        {
            /* Get image */
            img = cvQueryFrame(cap);
            if (img != NULL)
            {
                /* Convert gray image */
                cvCvtColor(img, g_img, CV_RGB2GRAY);

                /* Write image. */
                c_WriteAVI.WriteImage((BYTE*)img->imageData);

                /* Show image in window. */
                cvShowImage("Test", img);
            }

            if ((GetTickCount() - time)>15000)
            {
                printf("Time out\n");
                time = GetTickCount();
                c_WriteAVI.ReleaseVideo();
                /* Create video */
                c_WriteAVI.CreateVideo(L"output.wmv",
                    MFVideoFormat_WMV3,
                    800000,
                    640,
                    480,
                    30,
                    3);
            }

            int key = cvWaitKey(15);
            if ((char)key == 'q')
            {
                break;
            }
        }

        /* Release image */
        cvReleaseImage(&g_img);

        /* Release video */
        c_WriteAVI.ReleaseVideo();

        /* Deallocate memory */
        cvReleaseCapture(&cap);
        //CoUninitialize();
    }
}