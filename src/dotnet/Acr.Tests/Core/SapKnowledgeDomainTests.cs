using Acr.Core.Models;
using Acr.Core.SapKnowledge;

namespace Acr.Tests.Core;

public sealed class SapKnowledgeDomainTests
{
    [Fact]
    public void EpistemicTier_CautionAndAuthorityRules_EnforceUnverifiedHypothesisGuardrail()
    {
        Assert.True(EpistemicTierRules.IsAuthoritative(EpistemicTier.CompanyVerified));
        Assert.True(EpistemicTierRules.IsAuthoritative(EpistemicTier.EnvironmentValidated));
        Assert.False(EpistemicTierRules.IsAuthoritative(EpistemicTier.OfficialSapDoc));
        Assert.False(EpistemicTierRules.IsAuthoritative(EpistemicTier.GeneralSapReference));
        Assert.False(EpistemicTierRules.IsAuthoritative(EpistemicTier.UnverifiedHypothesis));

        Assert.True(EpistemicTierRules.RequiresCautionBanner(EpistemicTier.UnverifiedHypothesis, FragmentLifecycleState.Reviewed));
        Assert.True(EpistemicTierRules.RequiresCautionBanner(EpistemicTier.GeneralSapReference, FragmentLifecycleState.Reviewed));
        Assert.True(EpistemicTierRules.RequiresCautionBanner(EpistemicTier.OfficialSapDoc, FragmentLifecycleState.Reviewed));
        Assert.False(EpistemicTierRules.RequiresCautionBanner(EpistemicTier.OfficialSapDoc, FragmentLifecycleState.EnvironmentValidated));
        Assert.False(EpistemicTierRules.RequiresCautionBanner(EpistemicTier.CompanyVerified, FragmentLifecycleState.Reviewed));
    }

    [Fact]
    public void FragmentLifecycleRules_AllowsValidTransitions_AndRejectsInvalidJumps()
    {
        Assert.True(FragmentLifecycleRules.CanTransition(FragmentLifecycleState.Imported, FragmentLifecycleState.Extracted));
        Assert.True(FragmentLifecycleRules.CanTransition(FragmentLifecycleState.Extracted, FragmentLifecycleState.Reviewed));
        Assert.True(FragmentLifecycleRules.CanTransition(FragmentLifecycleState.Reviewed, FragmentLifecycleState.EnvironmentValidated));
        Assert.True(FragmentLifecycleRules.CanTransition(FragmentLifecycleState.EnvironmentValidated, FragmentLifecycleState.Deprecated));
        Assert.True(FragmentLifecycleRules.CanTransition(FragmentLifecycleState.Deprecated, FragmentLifecycleState.Reviewed));

        Assert.False(FragmentLifecycleRules.CanTransition(FragmentLifecycleState.Imported, FragmentLifecycleState.EnvironmentValidated));
        Assert.False(FragmentLifecycleRules.CanTransition(FragmentLifecycleState.Extracted, FragmentLifecycleState.EnvironmentValidated));
        Assert.False(FragmentLifecycleRules.CanTransition(FragmentLifecycleState.Deprecated, FragmentLifecycleState.EnvironmentValidated));
        Assert.False(FragmentLifecycleRules.CanTransition(FragmentLifecycleState.Reviewed, FragmentLifecycleState.Reviewed));

        Assert.Throws<InvalidOperationException>(() =>
            FragmentLifecycleRules.ValidateTransition(FragmentLifecycleState.Imported, FragmentLifecycleState.EnvironmentValidated));
    }

    [Fact]
    public void SapTransactionCatalogRules_ClassifiesEinaMastStpoAsSapTable_AndZDescargaAsZTransaction()
    {
        List<TransactionEntry> catalog = SapTransactionCatalogRules.GetRequiredInitialCodes()
            .Select(code => new TransactionEntry
            {
                Code = code,
                Title = $"Entry {code}",
                Purpose = $"Purpose for {code}",
                ObjectKind = code is "EINA" or "MAST" or "STPO"
                    ? SapObjectKind.SapTable
                    : (code == "ZDescarga" ? SapObjectKind.ZTransaction : SapObjectKind.StandardTransaction)
            })
            .ToList();

        Assert.Empty(SapTransactionCatalogRules.ValidateCatalog(catalog));

        var mislabeledEina = new TransactionEntry
        {
            Code = "EINA",
            Title = "EINA",
            Purpose = "Vendor Info Record",
            ObjectKind = SapObjectKind.StandardTransaction
        };
        Assert.NotEmpty(SapTransactionCatalogRules.ValidateEntry(mislabeledEina));

        var mislabeledZDescarga = new TransactionEntry
        {
            Code = "ZDescarga",
            Title = "ZDescarga",
            Purpose = "Dock unload",
            ObjectKind = SapObjectKind.StandardTransaction
        };
        Assert.NotEmpty(SapTransactionCatalogRules.ValidateEntry(mislabeledZDescarga));

        var unsafeVkp2 = new TransactionEntry
        {
            Code = "VKP2",
            Title = "VKP2",
            Purpose = "Sales price lookup",
            ObjectKind = SapObjectKind.StandardTransaction
        };
        Assert.NotEmpty(SapTransactionCatalogRules.ValidateEntry(unsafeVkp2));

        TransactionEntry safeVkp2 = unsafeVkp2 with
        {
            OperationalRisks = ["VKP2 shows retail selling price in CRC, never import unit cost (EKPO-NETPR)."]
        };
        Assert.Empty(SapTransactionCatalogRules.ValidateEntry(safeVkp2));
    }

    [Fact]
    public void SapKnowledgeEntry_SeparatesSapGuiVersionFromBackendRelease_AndPreserves15Fields()
    {
        var entry = new SapKnowledgeEntry
        {
            EntryId = "sap.tcode.me80fn",
            Title = "ME80FN — Purchasing General Analyses",
            Description = "Mass lookup of PO 43 and Pedido 41 quantities and net import price.",
            TransactionCodes = ["ME80FN"],
            RelatedTransactions = ["ME23N", "MB51"],
            SapModule = "MM-PUR",
            FunctionalArea = "Purchasing & Receiving",
            EnvironmentInfo = new SapEnvironmentInfo
            {
                SapGuiVersion = "SAP GUI 7.00",
                CompanyCode = "IM 01",
                PlantCode = "C001",
                ApplicabilityNotes = "Backend release pending System -> Status verification."
            },
            RequiredInputs = ["Purchase Order 43 or Pedido 41"],
            ExpectedOutputs = ["Order lines, UMP (PRK/UN), EKPO-NETPR"],
            NavigationInstructions = ["SAP Easy Access -> ME80FN"],
            ProcedureSteps = ["Enter PO list -> Execute -> Export spreadsheet"],
            RelevantTablesAndFields = [new SapTableFieldReference { TableName = "EKPO", FieldNames = ["EBELN", "NETPR"] }],
            DependenciesAndPrerequisites = ["PO 43 created"],
            OperationMode = SapOperationMode.ReadOnly,
            AuthorizationRequirements = ["ME80FN display authorization"],
            ConsequencesAndReversals = ["Read-only; no SAP records modified."],
            VerificationSteps = ["Compare exported unit sum against Packing List."],
            SourceProvenances = [new SourceProvenance { SourceDocument = "SOP-LOG-C001-01.docx", SourceFormat = "DOCX", PageOrSheet = "p. 4" }],
            LifecycleState = FragmentLifecycleState.Reviewed,
            EpistemicTier = EpistemicTier.CompanyVerified,
            Category = SapKnowledgeCategory.TransactionReferenceGuide,
            ObjectKind = SapObjectKind.StandardTransaction
        };

        Assert.Equal("SAP GUI 7.00", entry.EnvironmentInfo.SapGuiVersion);
        Assert.Equal(SapBackendRelease.UnknownPendingVerification, entry.EnvironmentInfo.BackendRelease);
        Assert.Equal(SapOperationMode.ReadOnly, entry.OperationMode);
        Assert.Single(entry.SourceProvenances);
        Assert.Single(entry.RelevantTablesAndFields);
    }

    [Fact]
    public void DiagnosticGuide_ValidatesAll10RequiredSections()
    {
        DiagnosticGuide complete = BuildCompleteGuide();
        Assert.True(complete.HasAllRequiredSections());
        Assert.Empty(complete.Validate());

        DiagnosticGuide missingSection = complete with { VerificationSteps = Array.Empty<string>() };
        Assert.False(missingSection.HasAllRequiredSections());
        Assert.Contains(missingSection.Validate(), e => e.Contains("Section 9", StringComparison.Ordinal));

        DiagnosticGuide badHypothesis = complete with
        {
            PossibleCauses = [new DiagnosticPossibleCause("Guess", EpistemicTier.UnverifiedHypothesis, IsConfirmedFact: true)]
        };
        Assert.False(badHypothesis.HasAllRequiredSections());
    }

    [Fact]
    public void DiagnosticGuide_RejectsNonAuthoritativeConfirmedFact_AndBlankListItems()
    {
        DiagnosticGuide complete = BuildCompleteGuide();

        DiagnosticGuide generalFact = complete with
        {
            PossibleCauses = [new DiagnosticPossibleCause("Forum note", EpistemicTier.GeneralSapReference, IsConfirmedFact: true)]
        };
        Assert.False(generalFact.HasAllRequiredSections());
        Assert.Contains(generalFact.Validate(), e => e.Contains("Section 6", StringComparison.Ordinal));

        DiagnosticGuide officialFact = complete with
        {
            PossibleCauses = [new DiagnosticPossibleCause("SAP Help note", EpistemicTier.OfficialSapDoc, IsConfirmedFact: true)]
        };
        Assert.False(officialFact.HasAllRequiredSections());

        DiagnosticGuide blankStep = complete with { InvestigationSteps = ["   "] };
        Assert.False(blankStep.HasAllRequiredSections());
        Assert.Contains(blankStep.Validate(), e => e.Contains("Section 4", StringComparison.Ordinal));

        DiagnosticGuide blankInfo = complete with { RequiredInvestigationInfo = ["Container ID", ""] };
        Assert.False(blankInfo.HasAllRequiredSections());

        DiagnosticGuide blankApproval = complete with { RequiredApprovalsAndValidation = ["\t"] };
        Assert.False(blankApproval.HasAllRequiredSections());

        DiagnosticGuide blankVerification = complete with { VerificationSteps = [" "] };
        Assert.False(blankVerification.HasAllRequiredSections());
    }

    [Fact]
    public void TransactionEntry_RequiresCompanyBehaviorCaution_ReflectsStandardSapAndValidation()
    {
        var unvalidatedStandard = new TransactionEntry
        {
            Code = "ME80FN",
            Title = "ME80FN",
            Purpose = "PO analysis",
            ObjectKind = SapObjectKind.StandardTransaction,
            EpistemicTier = EpistemicTier.OfficialSapDoc
        };
        Assert.True(unvalidatedStandard.RequiresCompanyBehaviorCaution);

        TransactionEntry validatedStandard = unvalidatedStandard with { EpistemicTier = EpistemicTier.EnvironmentValidated };
        Assert.False(validatedStandard.RequiresCompanyBehaviorCaution);

        var customZ = unvalidatedStandard with { Code = "ZDescarga", ObjectKind = SapObjectKind.ZTransaction };
        Assert.False(customZ.RequiresCompanyBehaviorCaution);
    }

    [Fact]
    public void DiagnosticContext_NeverInventsMissingIdentifiers()
    {
        var context = new DiagnosticContext
        {
            ContainerId = " IMP-104 ",
            Po43 = "   ",
            Pedido41 = null,
            Hu = "",
            MaterialCode = "1465571",
            DiscrepancyType = DiscrepancyType.Shortage,
            ShortageUnits = 24,
            SourceOrigin = DiagnosticContextOrigin.Reconciliation
        };

        Assert.Equal("IMP-104", context.ContainerId);
        Assert.Null(context.Po43);
        Assert.Null(context.Pedido41);
        Assert.Null(context.Hu);

        IReadOnlyDictionary<string, string> provided = context.GetProvidedIdentifiers();
        Assert.Equal(2, provided.Count);
        Assert.Equal("IMP-104", provided["ContainerId"]);
        Assert.Equal("1465571", provided["MaterialCode"]);

        IReadOnlyList<string> inferred = context.GetInferredBusinessObjects();
        Assert.Equal(["MATERIAL", "CONTENEDOR"], inferred);
        Assert.DoesNotContain("PO_43", inferred);
        Assert.DoesNotContain("PEDIDO_41", inferred);
        Assert.DoesNotContain("HU", inferred);
    }

    [Fact]
    public void InvestigationRecord_SeparatesObservedSuspectedConfirmed_AndGatesKbPromotion()
    {
        var ts = new DateTimeOffset(2026, 11, 18, 11, 0, 0, TimeSpan.Zero);
        var provenance = new SourceProvenance { SourceDocument = "MB51 Extract IMP-104", SourceFormat = "PersonalNote" };

        var record = new InvestigationRecord
        {
            RecordId = "INV-001",
            CreatedUtc = ts,
            UpdatedUtc = ts,
            Context = new DiagnosticContext { ContainerId = "IMP-104", MaterialCode = "1465571", DiscrepancyType = DiscrepancyType.Shortage },
            Findings =
            [
                new InvestigationFinding("F1", "Shortage of 24u on QC vs Pedido 41.", RootCauseState.ObservedDiscrepancy, provenance, ts),
                new InvestigationFinding("F2", "Possible unposted Mov 102 reversal.", RootCauseState.SuspectedCause, null, ts),
                new InvestigationFinding("F3", "Confirmed missing carton in EWM CLARIFICACION bin.", RootCauseState.ConfirmedRootCause, provenance, ts)
            ],
            ExternalActionsTaken = ["Verified HU in /SCWM/MON outside ACR."],
            VerificationResult = InvestigationVerificationStatus.VerifiedResolved,
            VerificationNotes = "Confirmed physical bin count matches corrected Mov 101."
        };

        Assert.Single(record.ObservedDiscrepancies);
        Assert.Single(record.SuspectedCauses);
        Assert.Single(record.ConfirmedRootCauses);
        Assert.Equal(RootCauseState.ConfirmedRootCause, record.CurrentRootCauseState);

        Assert.False(record.IsEligibleForPersonalKbPromotion(explicitUserConfirmation: false));
        Assert.Throws<InvalidOperationException>(() =>
            record.PromoteToPersonalKnowledgeEntry("note-1", "Title", explicitUserConfirmation: false, isEnvironmentValidated: false));

        InvestigationRecord unconfirmed = record with
        {
            Findings = [record.ObservedDiscrepancies[0], record.SuspectedCauses[0]]
        };
        Assert.False(unconfirmed.IsEligibleForPersonalKbPromotion(explicitUserConfirmation: true));

        InvestigationRecord unverified = record with { VerificationResult = InvestigationVerificationStatus.Pending };
        Assert.False(unverified.IsEligibleForPersonalKbPromotion(explicitUserConfirmation: true));

        Assert.True(record.IsEligibleForPersonalKbPromotion(explicitUserConfirmation: true));
        SapKnowledgeEntry promoted = record.PromoteToPersonalKnowledgeEntry(
            "kb.personal.inv001", "CLARIFICACION Short-Count Check", explicitUserConfirmation: true, isEnvironmentValidated: true);

        Assert.Equal(SapKnowledgeCategory.PersonalOperationalNote, promoted.Category);
        Assert.Equal(EpistemicTier.EnvironmentValidated, promoted.EpistemicTier);
        Assert.Equal(FragmentLifecycleState.EnvironmentValidated, promoted.LifecycleState);
        Assert.Single(promoted.SourceProvenances);
    }

    [Fact]
    public void InvestigationRecord_IsEligibleForPersonalKbPromotion_ReturnsFalseWhenProvenanceMissing()
    {
        var ts = new DateTimeOffset(2026, 11, 18, 11, 0, 0, TimeSpan.Zero);
        var noProvenanceRecord = new InvestigationRecord
        {
            RecordId = "INV-NOPROV",
            CreatedUtc = ts,
            UpdatedUtc = ts,
            Context = new DiagnosticContext { ContainerId = "IMP-104", DiscrepancyType = DiscrepancyType.Shortage },
            Findings = [new InvestigationFinding("F1", "Confirmed root cause without source.", RootCauseState.ConfirmedRootCause, null, ts)],
            VerificationResult = InvestigationVerificationStatus.VerifiedResolved
        };

        Assert.False(noProvenanceRecord.IsEligibleForPersonalKbPromotion(explicitUserConfirmation: true));
        Assert.Throws<InvalidOperationException>(() =>
            noProvenanceRecord.PromoteToPersonalKnowledgeEntry("kb.noprov", "Title", explicitUserConfirmation: true, isEnvironmentValidated: false));
    }

    [Fact]
    public void DeterministicRanker_ExcludesDeprecatedByDefault_AndRanksTCodeAboveBody()
    {
        IReadOnlyList<SapKnowledgeEntry> docs = BuildSampleRankerEntries();

        IReadOnlyList<RankedSapResult<SapKnowledgeEntry>> ranked = DeterministicRanker.Rank(
            docs, new SapSearchRequest { QueryText = "ME80FN faltante" });

        Assert.Equal(2, ranked.Count);
        Assert.DoesNotContain(ranked, r => r.Document.EntryId == "doc.deprecated");
        Assert.Equal("doc.tcode", ranked[0].Document.EntryId);
        Assert.True(ranked[0].RelevanceScore > ranked[1].RelevanceScore);
        Assert.Contains(ranked[0].MatchedReasons, r => r.StartsWith("SynonymExpansion:faltante->shortage", StringComparison.Ordinal));
        Assert.False(ranked[0].RequiresCautionBanner);
        Assert.True(ranked[1].RequiresCautionBanner);
    }

    [Fact]
    public void DeterministicRanker_EnforcesZeroRelevanceGuard_AndDeterministicRepeatability()
    {
        IReadOnlyList<SapKnowledgeEntry> docs = BuildSampleRankerEntries();

        Assert.Empty(DeterministicRanker.Rank([docs[2]], new SapSearchRequest { QueryText = "zzzzunknownquery" }));

        IReadOnlyList<RankedSapResult<SapKnowledgeEntry>> first = DeterministicRanker.Rank(
            docs, new SapSearchRequest { QueryText = "ME80FN faltante" });
        IReadOnlyList<RankedSapResult<SapKnowledgeEntry>> second = DeterministicRanker.Rank(
            docs, new SapSearchRequest { QueryText = "ME80FN faltante" });

        Assert.Equal(first.Select(r => r.Document.EntryId), second.Select(r => r.Document.EntryId));
        Assert.Equal(first.Select(r => r.TotalScore), second.Select(r => r.TotalScore));
    }

    [Fact]
    public void DeterministicRanker_SynonymExpansion_DoesNotLeakSubstrings_AndMatchesPhrases()
    {
        var mov641 = BuildEntry("doc.mov641", "Despacho Tiendas Mov 641", "Transfer order");
        var ewm41461 = BuildEntry("doc.ewm", "EWM Task 41461", "Warehouse task");
        var imp401 = BuildEntry("doc.imp401", "Container IMP-401 Check", "Dock inspection");
        var pedido41Doc = BuildEntry("doc.p41", "Ingreso de Pedido 41", "Nacionalizacion en SAP");
        var ncrDoc = BuildEntry("doc.ncr40", "Nota de Credito para proveedor", "Reclamo NCR 40");

        IReadOnlyList<RankedSapResult<SapKnowledgeEntry>> nacResults = DeterministicRanker.Rank(
            [mov641, ewm41461, pedido41Doc], new SapSearchRequest { QueryText = "nacionalizacion" });
        Assert.Single(nacResults);
        Assert.Equal("doc.p41", nacResults[0].Document.EntryId);

        IReadOnlyList<RankedSapResult<SapKnowledgeEntry>> ncrResults = DeterministicRanker.Rank(
            [imp401, ncrDoc], new SapSearchRequest { QueryText = "ncr" });
        Assert.Single(ncrResults);
        Assert.Equal("doc.ncr40", ncrResults[0].Document.EntryId);
    }

    [Fact]
    public void DeterministicRanker_MatchesUnderscoreBusinessObjects_AndMultiWordTransactionCodes()
    {
        var poDoc = BuildEntry("doc.po43", "Orden de Compra", "Revision general") with { BusinessObjects = ["PO_43"] };
        var zFacturasDoc = BuildEntry("doc.zfac", "Facturacion", "Reporte local") with { TransactionCodes = ["Z Facturas"] };
        var z51Doc = BuildEntry("doc.z51", "Liquidacion", "Proceso aduana") with { TransactionCodes = ["Z51 -Z52"] };

        IReadOnlyList<RankedSapResult<SapKnowledgeEntry>> poRes = DeterministicRanker.Rank(
            [poDoc], new SapSearchRequest { QueryText = "PO_43" });
        Assert.Single(poRes);
        Assert.Contains(poRes[0].MatchedReasons, r => r == "BusinessObject:po_43");

        IReadOnlyList<RankedSapResult<SapKnowledgeEntry>> zFacRes = DeterministicRanker.Rank(
            [zFacturasDoc], new SapSearchRequest { QueryText = "Z Facturas" });
        Assert.Single(zFacRes);
        Assert.Contains(zFacRes[0].MatchedReasons, r => r == "TCode:z facturas");

        IReadOnlyList<RankedSapResult<SapKnowledgeEntry>> z51Res = DeterministicRanker.Rank(
            [z51Doc], new SapSearchRequest { QueryText = "Z51 -Z52" });
        Assert.Single(z51Res);
        Assert.Contains(z51Res[0].MatchedReasons, r => r == "TCode:z51 -z52");
    }

    [Fact]
    public void DeterministicRanker_TieBreaking_Filters_IncludeDeprecated_AndEmptyQuery()
    {
        var docB = BuildEntry("doc.b", "MIGO Guide", "Receipt", EpistemicTier.EnvironmentValidated, FragmentLifecycleState.Reviewed, SapKnowledgeCategory.OperationalProcedure, SapObjectKind.StandardTransaction);
        var docA = BuildEntry("doc.a", "MIGO Guide", "Receipt", EpistemicTier.EnvironmentValidated, FragmentLifecycleState.Reviewed, SapKnowledgeCategory.OperationalProcedure, SapObjectKind.StandardTransaction);
        var docHigherState = BuildEntry("doc.z", "MIGO Guide", "Receipt", EpistemicTier.EnvironmentValidated, FragmentLifecycleState.EnvironmentValidated, SapKnowledgeCategory.TransactionReferenceGuide, SapObjectKind.ZTransaction);
        var docDeprecated = BuildEntry("doc.dep", "MIGO Old", "Receipt", EpistemicTier.CompanyVerified, FragmentLifecycleState.Deprecated, SapKnowledgeCategory.OperationalProcedure, SapObjectKind.StandardTransaction);

        IReadOnlyList<RankedSapResult<SapKnowledgeEntry>> emptyQuery = DeterministicRanker.Rank(
            [docB, docA, docHigherState, docDeprecated], new SapSearchRequest { QueryText = null, Context = null });
        Assert.Equal(["doc.z", "doc.a", "doc.b"], emptyQuery.Select(r => r.Document.EntryId));

        IReadOnlyList<RankedSapResult<SapKnowledgeEntry>> withDeprecated = DeterministicRanker.Rank(
            [docB, docDeprecated], new SapSearchRequest { IncludeDeprecated = true });
        Assert.Equal(["doc.b", "doc.dep"], withDeprecated.Select(r => r.Document.EntryId));

        IReadOnlyList<RankedSapResult<SapKnowledgeEntry>> filtered = DeterministicRanker.Rank(
            [docB, docHigherState],
            new SapSearchRequest
            {
                CategoryFilter = SapKnowledgeCategory.TransactionReferenceGuide,
                ObjectKindFilter = SapObjectKind.ZTransaction,
                MinEpistemicTier = EpistemicTier.EnvironmentValidated
            });
        Assert.Single(filtered);
        Assert.Equal("doc.z", filtered[0].Document.EntryId);
    }

    private static SapKnowledgeEntry BuildEntry(
        string id,
        string title,
        string description,
        EpistemicTier tier = EpistemicTier.CompanyVerified,
        FragmentLifecycleState state = FragmentLifecycleState.Reviewed,
        SapKnowledgeCategory category = SapKnowledgeCategory.OperationalProcedure,
        SapObjectKind kind = SapObjectKind.StandardTransaction) => new()
    {
        EntryId = id,
        Title = title,
        Description = description,
        EpistemicTier = tier,
        LifecycleState = state,
        Category = category,
        ObjectKind = kind
    };

    private static IReadOnlyList<SapKnowledgeEntry> BuildSampleRankerEntries() =>
    [
        new SapKnowledgeEntry
        {
            EntryId = "doc.deprecated",
            Title = "ME80FN Legacy Guide",
            Description = "ME80FN old notes",
            TransactionCodes = ["ME80FN"],
            EpistemicTier = EpistemicTier.CompanyVerified,
            LifecycleState = FragmentLifecycleState.Deprecated
        },
        new SapKnowledgeEntry
        {
            EntryId = "doc.body",
            Title = "General Purchasing Notes",
            Description = "Miscellaneous checklist",
            ProcedureSteps = ["Check ME80FN for details"],
            EpistemicTier = EpistemicTier.GeneralSapReference,
            LifecycleState = FragmentLifecycleState.Reviewed
        },
        new SapKnowledgeEntry
        {
            EntryId = "doc.tcode",
            Title = "ME80FN Mass PO Analysis",
            Description = "Inspect shortages and PO 43 lines",
            TransactionCodes = ["ME80FN"],
            Keywords = ["shortage"],
            BusinessObjects = ["PO_43"],
            DiscrepancyTypes = [DiscrepancyType.Shortage],
            EpistemicTier = EpistemicTier.CompanyVerified,
            LifecycleState = FragmentLifecycleState.EnvironmentValidated
        }
    ];

    private static DiagnosticGuide BuildCompleteGuide() => new()
    {
        GuideId = "guide.shortage.mov101",
        Title = "Shortage Investigation (QC vs Mov 101)",
        TargetDiscrepancyTypes = [DiscrepancyType.Shortage],
        TargetBusinessObjects = ["PO_43", "PEDIDO_41", "MATERIAL"],
        TransactionCodes = ["MB51", "ME80FN"],
        ProblemSummaryAndAffectedObjects = new DiagnosticProblemSummary(
            "Physical QC units are lower than invoiced Pedido 41 / Mov 101.",
            ["PO_43", "PEDIDO_41", "MATERIAL"]),
        RequiredInvestigationInfo = ["Container ID", "Material Code", "Pedido 41"],
        RelevantTransactionsAndPaths =
        [
            new DiagnosticTransactionStep("MB51", SapObjectKind.StandardTransaction, "SAP Easy Access -> MB51", "Verify Mov 101/102")
        ],
        InvestigationSteps = ["1. Run MB51 for plant C001 and Mov 101-102.", "2. Cross-check against CONTROL DATA."],
        RelevantTablesAndRelationships =
        [
            new SapTableFieldReference { TableName = "MSEG", FieldNames = ["MATNR", "BWART", "MENGE"], RelationshipNote = "Movement line items" }
        ],
        PossibleCauses =
        [
            new DiagnosticPossibleCause("Unposted Mov 102 reversal or short-shipped vendor carton.", EpistemicTier.CompanyVerified, IsConfirmedFact: false)
        ],
        RecommendedResolutionsAndRisks =
        [
            new DiagnosticResolutionOption("Record shortage in Col N and prepare NCR 40...", "Verify before generating Credit Note.", IsReadOnlyGuidance: false)
        ],
        RequiredApprovalsAndValidation = ["Supervisor approval required for NCR 40..."],
        VerificationSteps = ["Confirm updated stock in MB52 and PARA N.C RECEP. total."],
        DocumentationReferences =
        [
            new DiagnosticCitation(
                new SourceProvenance { SourceDocument = "SOP-LOG-C001-01.docx", SourceFormat = "DOCX", PageOrSheet = "p. 9" },
                EpistemicTier.CompanyVerified,
                "Shortage NCR workflow")
        ]
    };
}
