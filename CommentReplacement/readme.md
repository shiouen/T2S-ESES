This project is used to remove French comments and the hashtag-surrounded indicators from .docx documents.

Usage:
Open the project in Visual Studio and place your .docx files in the T2S-ESES/CommentReplacement/Files folder.
Run the program in Debug and you will find the resulting documents in T2S-ESES/CommentReplacement/bin/Debug/Files.
In case you run it as Release you can find the resulting documents in T2S-ESES/CommentReplacement/bin/Release/Files.

The program catches most use cases and occurences that need to be removed, but not all of them.
After this program has been used one should compare the original document with the newly created 
one to make sure the file has been handled correctly.

Example:
<pre>
<O@@vyiwO8)d4fO2 Type="Contrainte">
<A@@Z20000000D60 Attribute="Nom court">
EA-MT-1710
</A@@Z20000000D60>
<T@@f10000000b20 Attribute="Commentaire">
Règle d'alimentation du type de sous compte

Pour MT042

Le champ 'Sub account type' est alimenté par la valeur par défaut 'L1' dans le message à enrichir.

## T2S-ESES-R3 #CRE# - Constraint modified for translation purpose only#

Feeding rule of the sub-account type
For MT042

The 'Sub account type' field is fed by the default value 'L1' in the message to be enriched.

##
</pre>

Result:
<pre>
<O@@vyiwO8)d4fO2 Type="Contrainte">
<A@@Z20000000D60 Attribute="Nom court">
EA-MT-1710
</A@@Z20000000D60>
<T@@f10000000b20 Attribute="Commentaire">

Feeding rule of the sub-account type
For MT042

The 'Sub account type' field is fed by the default value 'L1' in the message to be enriched.
</pre>
