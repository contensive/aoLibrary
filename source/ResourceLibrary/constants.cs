
namespace Contensive.Addons.ResourceLibrary {
    //
    public static class constants {
        //
        public const string spacer1x5 = "<img src=/ResourceLibrary/spacer.gif width=5 height=1>";
        public const string spacer1x10 = "<img src=/ResourceLibrary/spacer.gif width=10 height=1>";
        public const string spacer1x15 = "<img src=/ResourceLibrary/spacer.gif width=15 height=1>";
        public const string spacer1x20 = "<img src=/ResourceLibrary/spacer.gif width=20 height=1>";
        public const string spacer1x30 = "<img src=/ResourceLibrary/spacer.gif width=30 height=1>";
        public const string spacer1x50 = "<img src=/ResourceLibrary/spacer.gif width=50 height=1>";
        //
        // -- sample
        public const string ButtonCancel = "Cancel";
        public const string ButtonDelete = "Delete";
        public const string ButtonApply = "Apply";
        public const string ButtonClose = "Close";
        //
        // -- sample
        public const string RequestNameButton = "button";
        public const string RequestNameLibraryUpload = "LibraryUpload";
        public const string RequestNameLibraryName = "LibraryName";
        public const string RequestNameLibraryDescription = "LibraryDescription";
        public const string xxxxx = "xxxx";
        //
        // -- errors for resultErrList
        public enum resultErrorEnum {
            errPermission = 50,
            errDuplicate = 100,
            errVerification = 110,
            errRestriction = 120,
            errInput = 200,
            errAuthentication = 300,
            errAdd = 400,
            errSave = 500,
            errDelete = 600,
            errLookup = 700,
            errLoad = 710,
            errContent = 800,
            errMiscellaneous = 900
        }
        //
        // -- http errors
        public enum httpErrorEnum {
            badRequest = 400,
            unauthorized = 401,
            paymentRequired = 402,
            forbidden = 403,
            notFound = 404,
            methodNotAllowed = 405,
            notAcceptable = 406,
            proxyAuthenticationRequired = 407,
            requestTimeout = 408,
            conflict = 409,
            gone = 410,
            lengthRequired = 411,
            preconditionFailed = 412,
            payloadTooLarge = 413,
            urlTooLong = 414,
            unsupportedMediaType = 415,
            rangeNotSatisfiable = 416,
            expectationFailed = 417,
            teapot = 418,
            methodFailure = 420,
            // enhanceYourCalm = 420,
            misdirectedRequest = 421,
            unprocessableEntity = 422,
            locked = 423,
            failedDependency = 424,
            upgradeRequired = 426,
            preconditionRequired = 428,
            tooManyRequests = 429,
            requestHeaderFieldsTooLarge = 431,
            loginTimeout = 440,
            noResponse = 444,
            retryWith = 449,
            redirect = 451,
            // unavailableForLegalReasons = 451,
            sslCertificateError = 495,
            sslCertificateRequired = 496,
            httpRequestSentToSecurePort = 497,
            invalidToken = 498,
            clientClosedRequest = 499,
            // tokenRequired = 499,
            internalServerError = 500
        }
    }
}
