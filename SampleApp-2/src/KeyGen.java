import javax.crypto.Mac;
import javax.crypto.spec.SecretKeySpec;
import java.security.InvalidKeyException;
import java.util.Base64;
import java.util.Date;

public class KeyGen {
    public static final String HMAC_SHA256 = "HmacSHA256";

    public static void main( String[] args )
    {
        try {

            String secret_key="",secret_key_timestamp="";
            // Initializing key in some variable. You will receive this key from Eko via email
            String key = "7d29d27a-ea96-4952-96d8-28856885dcfc";
            // test - 7d29d27a-ea96-4952-96d8-28856885dcfc
            // saransh prod - 9ddcb40b-0dc2-40e5-a86b-a4d12414336e
            //24c94eae-046c-411d-9a7b-174e564a211b
            //d2fe1d99-6298-4af2-8cc5-d97dcf46df30
            //String key = "e2ff5243-2bca-4640-97da-2319250283e5";
            //String encodedKey = Base64.encodeBase64String(key.getBytes());
            String encodedKey = Base64.getEncoder().encodeToString(key.getBytes());
            // Encode it using base64

            // Get secret_key_timestamp: current timestamp in milliseconds since UNIX epoch as STRING
            // Check out https://currentmillis.com to understand the timestamp format
            Date date = new Date();
            secret_key_timestamp = Long.toString(date.getTime());

            // Computes the signature by hashing the salt with the encoded key
            Mac sha256_HMAC;

            sha256_HMAC = Mac.getInstance(HMAC_SHA256);

            SecretKeySpec signature = new SecretKeySpec(encodedKey.getBytes(), HMAC_SHA256);
            try {
                sha256_HMAC.init(signature);
            } catch (InvalidKeyException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }

            // Encode it using base64 to get secret-key
            //secret_key = Base64.encodeBase64String(sha256_HMAC.doFinal(secret_key_timestamp.getBytes())).trim();
            secret_key = Base64.getEncoder().encodeToString(sha256_HMAC.doFinal(secret_key_timestamp.getBytes())).trim();
            System.out.println("secret_key : "+secret_key);
            System.out.println("secret_key_timestamp : "+secret_key_timestamp);


        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
